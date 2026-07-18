// Records which PowerPoint Office.js features the current host supports, so the
// UI can warn or fall back. Detection is defensive: in a non-Office context
// (e.g. unit tests) every capability reports false instead of throwing.
export interface PowerPointCapabilities {
  supportsShapeTags: boolean; // PowerPointApi 1.3 — chart model persistence
  supportsGrouping: boolean; // PowerPointApi 1.8 — addGroup (one movable object)
  supportsSelectedShapes: boolean; // PowerPointApi 1.5 — getSelectedShapes
  supportsSelectionEvents: boolean; // DocumentSelectionChanged in PowerPoint
}

function isSet(set: string, version: string): boolean {
  try {
    return (
      typeof Office !== "undefined" &&
      !!Office.context?.requirements?.isSetSupported(set, version)
    );
  } catch {
    return false;
  }
}

export function detectCapabilities(): PowerPointCapabilities {
  return {
    supportsShapeTags: isSet("PowerPointApi", "1.3"),
    supportsGrouping: isSet("PowerPointApi", "1.8"),
    supportsSelectedShapes: isSet("PowerPointApi", "1.5"),
    // Selection events are host-dependent; gate on the same set that exposes selection.
    supportsSelectionEvents: isSet("PowerPointApi", "1.5"),
  };
}
