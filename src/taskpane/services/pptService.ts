import type { CsvRow, ShapeMapping } from "../types";

export function isPowerPointApiSupported(): boolean {
  // getSelectedShapes / getSelectedSlides require PowerPointApi 1.5+ (Microsoft Learn).
  try {
    if (typeof Office === "undefined" || !Office.context?.requirements) {
      return false;
    }
    return Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
  } catch {
    return false;
  }
}

/**
 * Timer starts before `fn()` runs so a synchronous hang inside `PowerPoint.run` still gets a timeout
 * once the event loop can process timers (does not help if the host blocks the JS thread forever).
 */
function withTimeout<T>(fn: () => Promise<T>, ms: number, timeoutMessage: string): Promise<T> {
  return new Promise((resolve, reject) => {
    const t = window.setTimeout(() => reject(new Error(timeoutMessage)), ms);
    Promise.resolve()
      .then(fn)
      .then(
        (v) => {
          window.clearTimeout(t);
          resolve(v);
        },
        (e) => {
          window.clearTimeout(t);
          reject(e);
        }
      );
  });
}

function ensureOfficeReady(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof Office === "undefined" || typeof Office.onReady !== "function") {
      resolve();
      return;
    }
    let done = false;
    const finish = () => {
      if (done) return;
      done = true;
      resolve();
    };
    Office.onReady(() => finish());
    // If onReady never fires (broken host), unblock after 10s so the UI can show errors instead of hanging.
    window.setTimeout(finish, 10_000);
  });
}

async function ensureSupported(): Promise<void> {
  await ensureOfficeReady();
  if (typeof Office === "undefined") {
    throw new Error("Office.js is not loaded. Use this add-in inside PowerPoint.");
  }
  if (!isPowerPointApiSupported()) {
    throw new Error("This host doesn't support the required PowerPoint JavaScript APIs (PowerPointApi 1.5+).");
  }
}

/**
 * Lists text-capable shapes on the **currently selected** template slide by walking `slide.shapes`.
 * We intentionally do **not** use `getSelectedShapes()` — it can hard-freeze the task pane WebView on Mac PowerPoint.
 *
 * @remarks getSelectedSlides is PowerPointApi 1.5.
 */
export async function listMappableShapesOnTemplateSlide(): Promise<{
  templateSlideId: string;
  shapes: Array<{ shapeId: string; shapeName: string }>;
}> {
  await ensureSupported();
  return withTimeout(
    () =>
      PowerPoint.run(async (context) => {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items/id");
        await context.sync();

        if (selectedSlides.items.length !== 1) {
          throw new Error(
            "Select exactly one slide in the thumbnail pane (your template slide). " +
              "Click a single slide so only that slide is selected."
          );
        }

        const templateSlideId = selectedSlides.items[0].id;
        const slide = context.presentation.slides.getItem(templateSlideId);
        const shapes = slide.shapes;
        shapes.load("items/id,items/name,type");
        await context.sync();

        const mappableTypes = new Set<string>([
          PowerPoint.ShapeType.textBox,
          PowerPoint.ShapeType.placeholder,
          PowerPoint.ShapeType.geometricShape,
          PowerPoint.ShapeType.callout,
          "TextBox",
          "Placeholder",
          "GeometricShape",
          "Callout"
        ]);

        const out: Array<{ shapeId: string; shapeName: string }> = [];
        let unnamed = 0;
        for (const sh of shapes.items) {
          const t = sh.type as string;
          if (!mappableTypes.has(t)) continue;

          unnamed += 1;
          const rawName = (sh.name && String(sh.name).trim()) || "";
          const shapeName = rawName || `Text shape ${unnamed}`;
          out.push({ shapeId: sh.id, shapeName });
        }

        return { templateSlideId, shapes: out };
      }),
    60_000,
    "Timed out while reading shapes from the slide. Select your template slide and try Refresh shapes again."
  );
}

/**
 * If all mappings target the same template slide, return that id (skip an extra Office round-trip on preview).
 */
export function getTemplateSlideIdFromMappings(mappings: ShapeMapping[]): string | null {
  if (mappings.length === 0) return null;
  const id = mappings[0].templateSlideId;
  return mappings.every((m) => m.templateSlideId === id) ? id : null;
}

export async function getTemplateSlideIdFromSelection(): Promise<string> {
  await ensureSupported();
  return await PowerPoint.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items/id");
    await context.sync();
    if (selectedSlides.items.length !== 1) {
      throw new Error("Please select exactly one slide (your template slide).");
    }
    return selectedSlides.items[0].id;
  });
}

/** Apply CSV row to mapped shapes on one slide, within an existing request context (no nested PowerPoint.run). */
async function applyRowToSlideInContext(
  context: PowerPoint.RequestContext,
  slideId: string,
  row: CsvRow,
  mappings: ShapeMapping[],
  skipBlankValues?: boolean
): Promise<{ updated: number; warnings: string[] }> {
  const slide = context.presentation.slides.getItem(slideId);
  slide.load("id");
  await context.sync();

  const warnings: string[] = [];
  let updated = 0;

  for (const m of mappings) {
    const value = row[m.columnName] ?? "";
    if (skipBlankValues && (!value || value.trim().length === 0)) continue;

    const shape = await tryGetShapeOnSlideByIdOrName(context, slide, m.shapeId, m.shapeName);
    if (!shape) {
      warnings.push(
        `Missing shape for mapping "${m.label ?? m.columnName}" (shapeId=${m.shapeId}${m.shapeName ? `, name="${m.shapeName}"` : ""}).`
      );
      continue;
    }
    shape.textFrame.textRange.text = value ?? "";
    updated += 1;
  }

  await context.sync();
  return { updated, warnings };
}

function normalizeShapeName(name: string | undefined | null): string {
  return (name ?? "").trim();
}

async function tryGetShapeOnSlideByIdOrName(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  shapeId: string,
  shapeName?: string
): Promise<PowerPoint.Shape | null> {
  const targetName = normalizeShapeName(shapeName);

  // Try by id first (works on the template slide; duplicated slides get new ids).
  try {
    const s = slide.shapes.getItem(shapeId);
    // Do NOT require hasText — layout placeholders are often empty until you fill them.
    s.load("id,name,textFrame");
    await context.sync();
    if (s.textFrame) return s;
  } catch {
    // Shape id not on this slide (e.g. after duplicate) — fall through to name match.
  }

  if (!targetName) return null;

  const shapes = slide.shapes;
  shapes.load("items/id,items/name");
  await context.sync();

  const match = shapes.items.find((it) => normalizeShapeName(it.name) === targetName);
  if (!match) return null;

  match.load("textFrame");
  await context.sync();
  // Placeholders from slide masters always have a text frame even when empty.
  return match.textFrame ? match : null;
}

export async function applyRowToMappingsOnSlide(params: {
  slideId: string;
  row: CsvRow;
  mappings: ShapeMapping[];
  skipBlankValues?: boolean;
}): Promise<{ updated: number; warnings: string[] }> {
  await ensureSupported();
  const { slideId, row, mappings, skipBlankValues } = params;

  return await PowerPoint.run(async (context) => {
    return await applyRowToSlideInContext(context, slideId, row, mappings, skipBlankValues);
  });
}

export async function duplicateTemplateSlideAndPopulate(params: {
  templateSlideId: string;
  rows: CsvRow[];
  mappings: ShapeMapping[];
  onProgress?: (msg: string) => void;
  skipBlankValues?: boolean;
}): Promise<{ created: number; warnings: string[] }> {
  await ensureSupported();
  const { templateSlideId, rows, mappings, onProgress, skipBlankValues } = params;

  if (rows.length === 0) {
    return { created: 0, warnings: [] };
  }

  onProgress?.("Exporting template…");
  await new Promise<void>((resolve) => {
    window.setTimeout(() => resolve(), 0);
  });

  // One short run to export the template — avoids one giant batch that freezes the task pane WebView.
  const templateBase64 = await PowerPoint.run(async (context) => {
    const allSlides = context.presentation.slides;
    allSlides.load("items/id");

    const templateSlide = context.presentation.slides.getItem(templateSlideId);
    templateSlide.load("id");

    await context.sync();

    const templateIndex = allSlides.items.findIndex((s) => s.id === templateSlideId);
    if (templateIndex < 0) {
      throw new Error("Template slide not found. Please re-select your template slide and try again.");
    }

    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items/id");
    await context.sync();

    if (!selectedSlides.items.some((s) => s.id === templateSlideId)) {
      context.presentation.setSelectedSlides([templateSlideId]);
      await context.sync();
    }

    const templateSelection = context.presentation.getSelectedSlides();
    templateSelection.load("items/id");
    const base64Result = templateSelection.exportAsBase64Presentation();
    await context.sync();
    const b64 = base64Result.value;
    if (!b64) {
      throw new Error("Could not export the template slide. Try selecting the template slide and run again.");
    }
    return b64;
  });

  const warnings: string[] = [];
  let created = 0;
  // Slides are inserted *after* targetSlideId; chain inserts after each new slide.
  let insertAfterSlideId = templateSlideId;

  for (let i = 0; i < rows.length; i++) {
    onProgress?.(`Creating slide ${i + 1} of ${rows.length}...`);
    // Yield so React can repaint progress (single long run blocks the UI thread).
    await new Promise<void>((resolve) => {
      window.setTimeout(() => resolve(), 0);
    });

    await PowerPoint.run(async (context) => {
      context.presentation.insertSlidesFromBase64(templateBase64, {
        formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
        targetSlideId: insertAfterSlideId
      });
      await context.sync();

      const slidesAfterInsert = context.presentation.slides;
      slidesAfterInsert.load("items/id");
      await context.sync();

      const afterIdx = slidesAfterInsert.items.findIndex((s) => s.id === insertAfterSlideId);
      if (afterIdx < 0 || afterIdx + 1 >= slidesAfterInsert.items.length) {
        throw new Error("Could not locate the newly inserted slide after insert.");
      }

      const insertedSlide = slidesAfterInsert.getItemAt(afterIdx + 1);
      insertedSlide.load("id");
      await context.sync();

      insertAfterSlideId = insertedSlide.id;

      const { warnings: w } = await applyRowToSlideInContext(
        context,
        insertedSlide.id,
        rows[i],
        mappings,
        skipBlankValues
      );
      warnings.push(...w.map((x) => `[Slide ${i + 1}] ${x}`));
      created += 1;
    });
  }

  onProgress?.(`Done. Created ${created} slides.`);
  return { created, warnings };
}

