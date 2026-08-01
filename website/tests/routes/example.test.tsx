import { render, screen } from "@testing-library/react";
import { createRoutesStub, RouterContextProvider } from "react-router";
import { describe, expect, it } from "vitest";
import { exampleContext } from "~/context";
import ExampleRoute, { loader } from "~/routes/example";

describe("/example Route", () => {
  it("renders the message the loader provides", async () => {
    const context = new RouterContextProvider();
    context.set(exampleContext, "Message from context");

    const Stub = createRoutesStub(
      [{ path: "/example", Component: ExampleRoute, loader, HydrateFallback: () => null }],
      context,
    );

    render(<Stub initialEntries={["/example"]} />);

    expect(await screen.findByRole("heading", { name: "Message from context" })).toBeInTheDocument();
  });
});
