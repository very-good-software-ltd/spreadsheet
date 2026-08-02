import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { ViewerPage } from "~/pages/viewer";

describe("ViewerPage", () => {
  it("prompts for a file before one is chosen", () => {
    render(<ViewerPage />);

    expect(screen.getByText("Drop an .xlsx here, or choose a file above.")).toBeInTheDocument();
    expect(screen.getByRole("link", { name: "GitHub" })).toBeInTheDocument();
  });
});
