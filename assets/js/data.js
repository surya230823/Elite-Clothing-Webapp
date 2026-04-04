/**
 * Mock datasets for the dashboard (orders, mail, alerts).
 * Replace with API/fetch logic when you connect a backend.
 */
const MOCK_COMPLETED = [
  {
    id: "ORD-2401",
    vendor: "Loomcraft Textiles",
    cloth: "Organic cotton twill",
    completed: "2026-03-28",
    days: 14,
    material: "Cotton",
    orderType: "Bulk",
  },
  {
    id: "ORD-2398",
    vendor: "Harbor Mills",
    cloth: "Poly-viscose blend",
    completed: "2026-03-25",
    days: 21,
    material: "Blend",
    orderType: "Sample-led",
  },
  {
    id: "ORD-2392",
    vendor: "Southwind Fabrics",
    cloth: "Linen stripe",
    completed: "2026-03-20",
    days: 18,
    material: "Linen",
    orderType: "Bulk",
  },
  {
    id: "ORD-2388",
    vendor: "Loomcraft Textiles",
    cloth: "Recycled polyester",
    completed: "2026-03-15",
    days: 12,
    material: "Polyester",
    orderType: "Rush",
  },
];

const MOCK_ONGOING = [
  {
    id: "REQ-2412",
    vendor: "Atlas Weaving",
    material: "Merino wool jersey",
    progress: "QUOTATION PROVIDED",
    received: "2026-03-22",
    eta: "2026-04-10",
    incharge: "M. Chen",
  },
  {
    id: "REQ-2410",
    vendor: "Brightline Co.",
    material: "Cotton poplin",
    progress: "SAMPLE CONFIRMATION",
    received: "2026-03-26",
    eta: "2026-04-05",
    incharge: "A. Okonkwo",
  },
  {
    id: "REQ-2407",
    vendor: "Harbor Mills",
    material: "Stretch denim",
    progress: "PRICE CONFIRMATION",
    received: "2026-03-18",
    eta: "2026-04-02",
    incharge: "M. Chen",
  },
  {
    id: "REQ-2403",
    vendor: "Southwind Fabrics",
    material: "Silk charmeuse",
    progress: "ENQUIRY / REACHOUT",
    received: "2026-04-01",
    eta: "2026-04-20",
    incharge: "J. Patel",
  },
];

const MOCK_RECEIVED = [
  {
    handler: "A. Okonkwo",
    vendor: "Brightline Co.",
    received: "2026-04-02",
    material: "Cotton sateen",
    quoted: 18.4,
    qty: 2400,
    eta: "2026-04-18",
  },
  {
    handler: "M. Chen",
    vendor: "Atlas Weaving",
    received: "2026-04-01",
    material: "Wool jersey",
    quoted: null,
    qty: 800,
    eta: "2026-04-12",
  },
  {
    handler: "J. Patel",
    vendor: "Southwind Fabrics",
    received: "2026-03-30",
    material: "Linen blend",
    quoted: 22.1,
    qty: 1200,
    eta: "2026-04-08",
  },
  {
    handler: "M. Chen",
    vendor: "Harbor Mills",
    received: "2026-03-28",
    material: "Poly twill",
    quoted: 14.9,
    qty: 5000,
    eta: "2026-04-22",
  },
];

const MOCK_AWAITING_ELITE = [
  {
    id: "REQ-2415",
    vendor: "Nova Threadworks",
    issue: "Approve revised stripe spec",
    stage: "SAMPLE CONFIRMATION",
    waiting: "elite",
  },
  {
    id: "REQ-2411",
    vendor: "Brightline Co.",
    issue: "Send final quotation PDF",
    stage: "QUOTATION PROVIDED",
    waiting: "elite",
  },
  {
    id: "REQ-2409",
    vendor: "Atlas Weaving",
    issue: "Confirm buffer % for rush order",
    stage: "PRICE CONFIRMATION",
    waiting: "elite",
  },
  {
    id: "REQ-2406",
    vendor: "Harbor Mills",
    issue: "Vendor uploaded counter-offer",
    stage: "PRICE CONFIRMATION",
    waiting: "vendor",
  },
];

const MOCK_ALERTS = [
  {
    id: "a1",
    level: "high",
    title: "REQ-2415 — sample lab results overdue",
    meta: "Due 2026-04-03 · Nova Threadworks",
  },
  {
    id: "a2",
    level: "med",
    title: "Pricing review for REQ-2410",
    meta: "Brightline Co. · Awaiting margin sign-off",
  },
  {
    id: "a3",
    level: "med",
    title: "Mail thread idle 5+ days",
    meta: "Southwind Fabrics · REQ-2403",
  },
  {
    id: "a4",
    level: "low",
    title: "Weekly pipeline digest ready",
    meta: "Export available from overview",
  },
];

const MOCK_MAIL = [
  {
    id: "m1",
    from: "buying@brightline.co",
    subject: "Stripe repeat and MOQ — cotton poplin",
    summary:
      "Brightline requests 1.5m width cotton poplin with 3 blue / 2 red / 1 yellow stripes. MOQ 3,000m. Need price per meter and lead time after sample approval.",
    date: "2026-04-02",
  },
  {
    id: "m2",
    from: "procurement@atlasmills.io",
    subject: "Re: Merino jersey — shrinkage assumptions",
    summary:
      "Atlas asks to confirm 3% shrinkage and whether buffer includes freight. They need updated quote by Friday.",
    date: "2026-04-01",
  },
  {
    id: "m3",
    from: "ops@southwindfabrics.com",
    subject: "Linen blend enquiry",
    summary:
      "New enquiry for linen-viscose 54\" width, natural dye compatible. Quantity TBD; need ballpark range first.",
    date: "2026-03-31",
  },
];
