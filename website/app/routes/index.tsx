import { LandingPage } from "~/pages/landing";

export function meta() {
  return [
    { title: "@very-good-software/spreadsheet" },
    {
      name: "description",
      content:
        "Read Excel .xlsx and OpenDocument .ods files in Node and the browser without holding the whole sheet in memory, and fill in .xlsx templates without disturbing the rest of the file.",
    },
  ];
}

export default function IndexRoute() {
  return <LandingPage />;
}
