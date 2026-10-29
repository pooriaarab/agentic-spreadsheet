import { useCallback, useEffect, useMemo, useState, type FormEvent } from "react";
import { AppShell, Button, Dialog, EmptyState, Input, Select, SidebarNav, useTheme, type ThemePreference } from "agentic-design/react";
import { SheetsApiError } from "../src/client.js";
import { clientFor, errorText, getToken, isUnauthorized, setToken } from "./session.js";
import { SheetView } from "./SheetView.js";

const bookFromHash = (): string | null => {
  const id = decodeURIComponent(location.hash.replace(/^#\//, ""));
  return id === "" ? null : id;
};

function SignIn({ onSignIn, notice }: { onSignIn: (token: string) => void; notice?: string }) {
  const [key, setKey] = useState("");
  const submit = (event: FormEvent) => {
    event.preventDefault();
    if (key.trim() !== "") onSignIn(key.trim());
  };
  return (
    <main className="center">
      <form className="card" onSubmit={submit}>
        <h1>Sign in to Sheets</h1>
        <p className="muted">Paste an API key. This tab keeps it until you close the tab.</p>
        {notice && <p role="alert" className="error-text">{notice}</p>}
        <Input label="API key" type="password" autoComplete="off" value={key} onChange={(e) => setKey(e.target.value)} />
        <Button type="submit">Sign in</Button>
      </form>
    </main>
  );
}

function ThemePick() {
  const { preference, setPreference } = useTheme();
  return (
    <Select label="Theme" hideLabel value={preference} onChange={(e) => setPreference(e.target.value as ThemePreference)}>
      <option value="system">Theme: system</option>
      <option value="light">Theme: light</option>
      <option value="dark">Theme: dark</option>
    </Select>
  );
}

export function App() {
  const [token, setTokenState] = useState(getToken);
  const [notice, setNotice] = useState<string>();
  const [book, setBook] = useState(bookFromHash);
  const [books, setBooks] = useState<{ status: "loading" } | { status: "error"; message: string } | { status: "ready"; list: string[] }>({ status: "loading" });
  const [naming, setNaming] = useState(false);
  const [name, setName] = useState("");
  const [nameError, setNameError] = useState<string>();
  const client = useMemo(() => (token === null ? null : clientFor(token)), [token]);

  const signOut = useCallback((why?: string) => {
    setToken(null);
    setTokenState(null);
    setNotice(why);
  }, []);

  useEffect(() => {
    const onHash = () => setBook(bookFromHash());
    addEventListener("hashchange", onHash);
    return () => removeEventListener("hashchange", onHash);
  }, []);

  const load = useCallback(async () => {
    if (client === null) return;
    setBooks({ status: "loading" });
    try {
      setBooks({ status: "ready", list: (await client.list()).books });
    } catch (error) {
      if (isUnauthorized(error)) return signOut("That key is not valid. Try another.");
      setBooks({ status: "error", message: errorText(error) });
    }
  }, [client, signOut]);
  useEffect(() => void load(), [load]);

  if (client === null || token === null) {
    return <SignIn notice={notice} onSignIn={(next) => (setToken(next), setTokenState(next))} />;
  }

  async function create(event: FormEvent) {
    event.preventDefault();
    try {
      const made = await client!.create({ book: name.trim() });
      setNaming(false);
      setName("");
      setNameError(undefined);
      await load();
      location.hash = `#/${encodeURIComponent(made.book)}`;
    } catch (error) {
      setNameError(error instanceof SheetsApiError && error.code === "conflict" ? `A workbook named "${name.trim()}" exists. Pick another name.` : errorText(error));
    }
  }

  const list = books.status === "ready" ? books.list : [];
  const open = () => {
    setNameError(undefined);
    setNaming(true);
  };
  return (
    <AppShell
      brand="Sheets"
      sidebar={<SidebarNav label="Workbooks" items={list.map((id) => ({ href: `#/${encodeURIComponent(id)}`, label: id, current: id === book }))} />}
      topbar={
        <>
          <Button size="sm" variant="secondary" onClick={open}>New workbook</Button>
          <ThemePick />
          <Button size="sm" variant="ghost" onClick={() => signOut()}>Sign out</Button>
        </>
      }
    >
      <Dialog
        open={naming}
        onOpenChange={setNaming}
        title="New workbook"
        description="Use letters, digits, dot, dash, or underscore."
      >
        <form className="stack" onSubmit={create}>
          <Input label="Workbook name" value={name} onChange={(e) => setName(e.target.value)} error={nameError} autoComplete="off" />
          <Button type="submit" disabled={name.trim() === ""}>Create workbook</Button>
        </form>
      </Dialog>
      {books.status === "loading" && <p role="status" className="muted">Loading workbooks…</p>}
      {books.status === "error" && (
        <div role="alert" className="banner">
          <p>Could not load the workbooks: {books.message}</p>
          <Button size="sm" variant="secondary" onClick={() => void load()}>Try again</Button>
        </div>
      )}
      {books.status === "ready" && book === null && (
        <section>
          <h1>Workbooks</h1>
          <EmptyState
            title={list.length === 0 ? "No workbooks yet" : "Pick a workbook"}
            description={list.length === 0 ? "Create a workbook to start a sheet." : "Choose a workbook in the list."}
            action={list.length === 0 ? <Button onClick={open}>New workbook</Button> : undefined}
          />
        </section>
      )}
      {book !== null && <SheetView key={book} client={client} book={book} token={token} onUnauthorized={() => signOut("That key is not valid. Try another.")} />}
    </AppShell>
  );
}
