#!/usr/bin/env node
import { VERSION } from "./version.js";

const arg = process.argv[2];
if (arg === "--version" || arg === "-V") {
  console.log(VERSION);
} else if (arg === undefined || arg === "--help" || arg === "-h") {
  console.log(`agentic-spreadsheet ${VERSION}\nUsage: agentic-spreadsheet [--help] [--version]`);
} else {
  console.error(`unknown command "${arg}"`);
  process.exit(2);
}
