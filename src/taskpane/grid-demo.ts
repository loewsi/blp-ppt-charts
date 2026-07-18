// Throwaway smoke test for the RevoGrid wrapper (not part of the add-in build target).
import { mountGrid, setGridData, getGridData } from "./grid";

const container = document.getElementById("grid") as HTMLElement;
mountGrid(container);
setGridData({
  type: "barColumn",
  categories: ["Q1", "Q2", "Q3", "Q4"],
  series: [
    { name: "Product A", color: "#2E75FF", values: [10, 12, 9, 14] },
    { name: "Product B", color: "#001C54", values: [6, 8, 7, 9] },
  ],
});

document.getElementById("dump")!.addEventListener("click", () => {
  document.getElementById("out")!.textContent = JSON.stringify(getGridData(), null, 2);
});
