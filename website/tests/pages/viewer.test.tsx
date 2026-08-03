import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { Demo } from "~/pages/viewer";

describe("Demo", () => {
  it("prompts for a file before one is chosen", () => {
    render(<Demo />);

    expect(screen.getByText("Drop an .xlsx here or choose a file above.")).toBeInTheDocument();
  });
});
