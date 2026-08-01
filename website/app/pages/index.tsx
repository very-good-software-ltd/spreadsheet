import { href, Link } from "react-router";

export function IndexPage() {
  return (
    <main className="flex h-screen items-center justify-center">
      <div className="flex flex-col items-center gap-16">
        <header className="flex flex-col items-center gap-9">
          <h1 className="leading text-2xl font-bold text-gray-800 dark:text-gray-100">New React Router App</h1>
        </header>
        <nav>
          <Link to={href("/example")}>Example</Link>
        </nav>
      </div>
    </main>
  );
}
