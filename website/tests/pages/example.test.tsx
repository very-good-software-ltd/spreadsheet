import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { ExamplePage } from "~/pages/example";

describe("ExamplePage", () => {
  it("displays the message it is given", () => {
    render(<ExamplePage message="The Message" />);

    expect(screen.getByRole("heading", { name: "The Message" })).toBeInTheDocument();
  });
});
