export type Offer = {
  source: string;
  unit_price: number;
  stock: number;
  lead_time_days: number;
  moq: number;
};

export type PricingTable = Record<string, Offer[]>;
export type XrefTable = Record<string, string[]>;

export type BomLine = {
  part_number: string;
  qty: number;
};

export type LineStatus = "ok" | "partial_stock" | "no_source";

export type PlanRow = {
  requested_pn: string;
  picked_pn: string;
  qty: number;
  status: LineStatus;
  substituted_from: string; // "" if no substitution
  sub_reason: string;
  best: Offer | null;
  alternates: Offer[];
  note: string;
  tried_parts: string[];
  ext_price: number; // unit_price * max(qty, moq), 0 if no best
};

export type OptimizeResult = {
  rows: PlanRow[];
  summary: {
    lines_total: number;
    fully_sourced: number;
    partial_stock: number;
    no_source: number;
    substituted: number;
    total_cost: number;
    critical_lead_days: number;
  };
};
