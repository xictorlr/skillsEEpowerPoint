import { afterEach, describe, expect, it, vi } from "vitest";
import { runCommand } from "../../src/office/runtime";

describe("runCommand", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("returns ok:true with the changed count on success", async () => {
    vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => Promise<number>) => fn({}) });

    const result = await runCommand("cmd", async () => 3);

    expect(result).toEqual({ ok: true, changed: 3 });
  });

  it("returns ok:false with the message of a plain thrown Error", async () => {
    vi.spyOn(console, "error").mockImplementation(() => undefined);
    vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => Promise<number>) => fn({}) });

    const result = await runCommand("cmd", async () => {
      throw new Error("x");
    });

    expect(result).toEqual({ ok: false, message: "x" });
  });

  it("maps a thrown ApiNotFound error code to the Spanish message", async () => {
    vi.spyOn(console, "error").mockImplementation(() => undefined);
    vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => Promise<number>) => fn({}) });

    const result = await runCommand("cmd", async () => {
      throw { code: "ApiNotFound", message: "m" };
    });

    expect(result).toEqual({
      ok: false,
      message: "Esta función requiere una versión más reciente de PowerPoint.",
    });
  });

  it("maps a thrown GeneralException error code to the Spanish message", async () => {
    vi.spyOn(console, "error").mockImplementation(() => undefined);
    vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => Promise<number>) => fn({}) });

    const result = await runCommand("cmd", async () => {
      throw { code: "GeneralException", message: "m" };
    });

    expect(result).toEqual({
      ok: false,
      message: "PowerPoint rechazó la operación. Puede que alguna forma esté bloqueada o no admita el cambio.",
    });
  });

  it("maps a thrown ItemNotFound error code to the Spanish message", async () => {
    vi.spyOn(console, "error").mockImplementation(() => undefined);
    vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => Promise<number>) => fn({}) });

    const result = await runCommand("cmd", async () => {
      throw { code: "ItemNotFound", message: "m" };
    });

    expect(result).toEqual({
      ok: false,
      message: "No se encontró la forma o diapositiva seleccionada.",
    });
  });

  it("falls back to 'code: message' for an unrecognized error code", async () => {
    vi.spyOn(console, "error").mockImplementation(() => undefined);
    vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => Promise<number>) => fn({}) });

    const result = await runCommand("cmd", async () => {
      throw { code: "Foo", message: "bar" };
    });

    expect(result).toEqual({ ok: false, message: "Foo: bar" });
  });
});
