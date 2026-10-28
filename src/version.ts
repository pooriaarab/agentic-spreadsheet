import pkg from "../package.json" with { type: "json" };

// The bundlers inline package.json, so the CLI, the library, and the Worker report one version.
export const VERSION: string = pkg.version;
