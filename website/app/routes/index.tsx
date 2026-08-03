import { LandingPage } from "~/pages/landing";

export function meta() {
  return [
    { title: "very-good-spreadsheet" },
    {
      name: "description",
      content: "Read Excel .xlsx files in Node and the browser, without loading the whole file into memory.",
    },
  ];
}

export default function IndexRoute() {
  return <LandingPage />;
}
