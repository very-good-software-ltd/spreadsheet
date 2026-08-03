import { LandingPage } from "~/pages/landing";

export function meta() {
  return [
    { title: "very-good-spreadsheet" },
    {
      name: "description",
      content: "Read Excel .xlsx files in Node and the browser, without holding the whole sheet in memory.",
    },
  ];
}

export default function IndexRoute() {
  return <LandingPage />;
}
