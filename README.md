# Parts Inventory

Mobile-first Next.js app for looking up appliance parts, bin locations, and stock levels against open service jobs.

## What’s included

- Standalone parts search by part number — no job has to be selected first
- Each part shows part number, bin location, quantity on hand, and a stock status badge
- Attach a part to a job from the part page or from a job card
- Linking a part never deducts quantity on hand; stock only changes if you receive or consume inventory separately
- Mock catalog of 5 common appliance parts

## Run locally

```bash
npm install
npm run dev -- --port 43127
```

Open [http://localhost:43127](http://localhost:43127).

## Stack

- Next.js App Router
- TypeScript
- Tailwind CSS
- Lucide React
