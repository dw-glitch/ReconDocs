import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Cruzamento de planilhas | ReconDocs",
  description: "Cruze quantas planilhas quiser, de qualquer formato, e gere um relatório Excel consolidado.",
};

export default function CrossCheckLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return children;
}
