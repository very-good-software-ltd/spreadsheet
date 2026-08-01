interface ExamplePageProps {
  message: string;
}

export function ExamplePage({ message }: ExamplePageProps) {
  return (
    <main className="flex h-screen items-center justify-center">
      <div className="flex flex-col items-center gap-16">
        <header className="flex flex-col items-center gap-9">
          <h1 className="leading text-2xl font-bold text-gray-800 dark:text-gray-100">{message}</h1>
        </header>
      </div>
    </main>
  );
}
