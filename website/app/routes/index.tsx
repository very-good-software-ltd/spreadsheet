import { ViewerPage } from "~/pages/viewer";

export function meta() {
  return [
    { title: "very-good-spreadsheet" },
    { name: "description", content: "Stream a large .xlsx into a virtualized table, in the browser." },
  ];
}

export default function IndexRoute() {
  return <ViewerPage />;
}
