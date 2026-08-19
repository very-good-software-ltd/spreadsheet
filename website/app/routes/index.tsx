import { LandingPage } from "~/pages/landing";

export function meta() {
  return [
    { title: "@very-good-software/spreadsheet" },
    {
      name: "description",
      content:
        "Read Excel .xlsx and OpenDocument .ods files in Node and the browser without holding the whole sheet in memory. Design a report in Excel and fill it from code with any number of rows, and the rest of the sheet stays as you designed it.",
    },
  ];
}

export default function IndexRoute() {
  return <LandingPage />;
}
