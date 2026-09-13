# Parts Inventory

Mobile-first Next.js app for looking up appliance parts, bin locations, and stock levels against open service jobs. Reads live from multiple inventory sources, including a SharePoint-hosted workbook, behind Microsoft sign-in.

## What's included

- Standalone parts search by part number — no job has to be selected first
- Each part shows part number, bin location, quantity on hand, and a stock status badge, derived from quantity so it can never drift out of sync (see `lib/status.ts`)
- Attach a part to a job from the part page or from a job card
- Linking a part never deducts quantity on hand; stock only changes if you receive or consume inventory separately
- Reads parts from multiple sources and merges them by part number (see `lib/getInventory.ts`, `lib/sources/`):
  - a bundled local catalog (always on, used for anything not covered by another source)
  - a SharePoint-hosted Excel workbook (optional — enabled once it's configured, see below)
- Every technician signs in with their own Microsoft account; the SharePoint source reads the workbook using that person's own delegated permissions, not a shared service account

## Run locally

```bash
npm install
cp .env.example .env.local   # fill in real values, see "Sign-in setup" below
npm run dev -- --port 43127
```

Open [http://localhost:43127](http://localhost:43127). The whole app requires signing in with Microsoft — see below before your first run.

## Sign-in setup (required)

The app is gated behind Microsoft sign-in (`proxy.ts`) because the SharePoint source needs a signed-in user's own Graph token to read anything. You need an Azure AD (Microsoft Entra ID) app registration:

1. In the Azure Portal, go to **Microsoft Entra ID → App registrations → New registration**.
2. Add a redirect URI (platform: **Web**): `{NEXTAUTH_URL}/api/auth/callback/azure-ad` — for local dev that's `http://localhost:43127/api/auth/callback/azure-ad`.
3. Generate a certificate and upload it (this app authenticates with a certificate, not a client secret — many tenants now block client secrets by policy, see "Why a certificate" below):
   ```bash
   npm install       # needed once, for the generator's dependency
   npm run generate-cert
   ```
   This writes `certs/azure-ad-private-key.pem` (keep it secret — it's already gitignored) and `certs/azure-ad-certificate.pem`, and prints two `AZURE_AD_CERT_..._BASE64` lines. Upload `certs/azure-ad-certificate.pem` under **Certificates & secrets → Certificates → Upload certificate**, and paste the two printed lines into `.env.local`.
4. Under **API permissions**, add the **delegated** Microsoft Graph permission `Sites.Selected` (plus `openid`, `profile`, `email`, `offline_access`) and grant admin consent.
   - `Sites.Selected` grants this app **no access to any SharePoint site by default** — deliberately, so it can't read or write anything outside the specific site(s) hosting inventory workbooks, unlike `Sites.Read.All`/`Sites.ReadWrite.All` which hand it access to every site the signed-in user can reach. There's no permission that scopes down to one *file* while leaving the rest of a site alone — a whole site is the finest unit Graph offers — but this keeps it to just the site(s) you explicitly grant, not the tenant.
   - Granting the permission in Entra ID is only half of it — you then have to grant this **specific app** access to **each SharePoint site** that has an inventory workbook, with the **Write** role (needed for updating/adding parts, not just reading). The easiest way is [PnP PowerShell](https://pnp.github.io/powershell/):
     ```powershell
     Install-Module -Name PnP.PowerShell -Scope CurrentUser   # once
     Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -Interactive
     Grant-PnPAzureADAppSitePermission -AppId "<AZURE_AD_CLIENT_ID>" -DisplayName "Parts Inventory" -Site "https://<tenant>.sharepoint.com/sites/<SiteName>" -Permissions Write
     ```
     Run `Grant-PnPAzureADAppSitePermission` once per site that has an inventory workbook — a source pointed at a site nobody has run this for will fail with a permissions error (Forbidden) even though sign-in itself works fine, since Sites.Selected alone grants nothing until a site is explicitly authorized this way.
     - To later revoke or list what's been granted: `Get-PnPAzureADAppSitePermission -Site "<url>"` / `Revoke-PnPAzureADAppSitePermission`.
     - No PowerShell available? The same grant can be made by calling `POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions` (in [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer), signed in as an admin) with body `{"roles": ["write"], "grantedToIdentities": [{"application": {"id": "<AZURE_AD_CLIENT_ID>", "displayName": "Parts Inventory"}}]}` — get `{site-id}` first from `GET /sites/{hostname}:{sitePath}`.
   - Changing this permission on an app that was already set up means everyone needs to sign out and back in afterward — an existing session's refresh token won't pick up the new scope on its own.
5. Copy the **Application (client) ID** and **Directory (tenant) ID** into `.env.local` as `AZURE_AD_CLIENT_ID` and `AZURE_AD_TENANT_ID`.
6. Set `NEXTAUTH_SECRET` to a random value (see `.env.example` for a PowerShell one-liner, or `openssl rand -base64 32` on macOS/Linux) and `NEXTAUTH_URL` to the URL the app runs at.

Without this, the app has nowhere to redirect a signed-out visitor and will 500 on every page — this step isn't optional.

### Why a certificate instead of a client secret

If creating a client secret in your app registration fails with *"Client secrets are blocked by a tenant-wide policy"*, that's a Microsoft Entra app management policy — increasingly the default — not something specific to this app. A certificate sidesteps it entirely and is what Microsoft recommends anyway. The certificate `npm run generate-cert` creates is self-signed, which is fine here: Azure AD only needs it to verify that whoever's calling the token endpoint holds the matching private key, not that it was issued by a public certificate authority.

Certificate-based auth to Azure AD's token endpoint has its own header requirements beyond generic OAuth (`PS256` + an `x5t#S256` header naming the certificate) that most generic libraries, including the one this app's sign-in library uses internally, don't produce correctly out of the box — see the comment at the top of `lib/auth/certificate.ts` for specifics if you're curious or need to debug a token-endpoint error.

The certificate is valid for 2 years. Before it expires, run `npm run generate-cert` again (after deleting the old files in `certs/`), upload the new certificate to Azure AD alongside the old one, update `.env.local`, and remove the old certificate from Azure AD once the new one is confirmed working.

## Connecting the SharePoint inventory source (optional)

Once sign-in works, point the app at the real workbook by setting these in `.env.local`:

```
SHAREPOINT_SITE_HOSTNAME=contoso.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/ServiceOps
SHAREPOINT_FILE_PATH=Shared Documents/Inventory.xlsx
SHAREPOINT_TABLE_NAME=Inventory
```

The workbook needs the inventory data formatted as a named **Excel Table** (select the range → Insert → Table → give it a name in the Table Design tab), not just a plain range — that's what makes reads robust to inserted rows and columns.

The column headers this reads are defined in `lib/sources/sharepoint-excel-source.ts` (`COLUMN_MAP`, near the top of the file) — edit that one object to match the real workbook's header text. `part_number` and `quantity_on_hand` are required; the source throws a clear error naming the headers it looked for if either is missing from the sheet, so a typo there is loud, not a silent empty catalog.

If the SharePoint source is unreachable (network issue, expired session, a renamed column) it's dropped for that request and a banner explains why — the rest of the catalog (local, and anything from other sources) still renders. It doesn't take the whole app down.

## Adding another source

Implement the `InventorySource` interface in `lib/sources/types.ts` (`fetchParts`/`fetchJobs`/`fetchJobParts`), add it to `buildSources()` in `lib/getInventory.ts`, and give it a priority position in that function's `priority` array — that's the order conflicting part numbers get resolved in. `lib/sources/local-source.ts` is the simplest reference implementation; `lib/sources/sharepoint-excel-source.ts` is a fuller one with error handling worth copying from.

## Run tests

```bash
npm run test
```

Covers status derivation, the cross-source merge/conflict logic, and the SharePoint row-mapping logic — all pure functions that don't need a live SharePoint connection.

## Stack

- Next.js App Router
- TypeScript
- Tailwind CSS
- Lucide React
- NextAuth.js (Microsoft Entra ID provider) for sign-in
- Microsoft Graph for reading the SharePoint-hosted workbook
