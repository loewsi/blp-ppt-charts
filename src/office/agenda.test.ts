import { describe, it, expect } from "vitest";
import { numberedList } from "./agenda";

describe("numberedList", () => {
  it("auto-numbers the chapters", () => {
    expect(numberedList(["Intro", "Market", "Plan"])).toBe("1.  Intro\n2.  Market\n3.  Plan");
  });

  it("handles a single chapter", () => {
    expect(numberedList(["Only"])).toBe("1.  Only");
  });
});
