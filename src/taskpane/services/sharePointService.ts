export interface IClauseItem {
  id: number;
  title: string;
  category: string;
  clauseText: string;
  versionLabel: string;
  approved: boolean;
}

export async function getClausesFromSharePoint(): Promise<IClauseItem[]> {
  const response = await fetch("http://localhost:3001/api/clauses", {
    method: "GET",
    headers: {
      Accept: "application/json"
    }
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`API request failed: ${response.status} ${response.statusText} - ${errorText}`);
  }

  const data = await response.json();

  return (data || []).map((item: any) => ({
    id: item.id ?? 0,
    title: item.title ?? "",
    category: item.category ?? "",
    clauseText: item.clauseText ?? "",
    versionLabel: item.versionLabel ?? "",
    approved: item.approved ?? false
  }));
}