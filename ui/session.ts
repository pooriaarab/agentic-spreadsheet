import { createClient, SheetsApiError, type SheetsClient } from "../src/client.js";

// The key lives in this tab only. A closed tab forgets it.
const KEY = "sheets.token";
export const getToken = (): string | null => sessionStorage.getItem(KEY);
export const setToken = (token: string | null): void => (token === null ? sessionStorage.removeItem(KEY) : sessionStorage.setItem(KEY, token));

export const clientFor = (token: string): SheetsClient => createClient({ baseUrl: location.origin, token });

export const errorText = (error: unknown): string => (error instanceof Error ? error.message : String(error));
export const isUnauthorized = (error: unknown): boolean => error instanceof SheetsApiError && error.code === "unauthorized";
