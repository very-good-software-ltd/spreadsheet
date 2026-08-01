import { render, screen } from "@testing-library/react";
import { MemoryRouter } from "react-router";
import { describe, expect, it } from "vitest";
import { IndexPage } from "~/pages";

describe("IndexPage", () => {
  it("shows a link to the example page", () => {
    render(
      <MemoryRouter>
        <IndexPage />
      </MemoryRouter>,
    );

    expect(screen.getByRole("link", { name: "Example" })).toHaveAttribute("href", "/example");
  });
});
