import { exampleContext } from "~/context";
import { ExamplePage } from "~/pages/example";
import type { Route } from "./+types/example";

export function meta() {
  return [{ title: "New React Router App" }, { name: "description", content: "Welcome to React Router!" }];
}

export function loader({ context }: Route.LoaderArgs) {
  return { message: context.get(exampleContext) };
}

export default function ExampleRoute({ loaderData }: Route.ComponentProps) {
  return <ExamplePage message={loaderData.message} />;
}
