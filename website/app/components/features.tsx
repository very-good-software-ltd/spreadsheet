import { FEATURES } from "~/site";

export function Features() {
  return (
    <section className="border-t border-gray-100 px-6 py-16 dark:border-gray-900">
      <div className="mx-auto max-w-5xl">
        <div className="grid gap-6 sm:grid-cols-2 lg:grid-cols-3">
          {FEATURES.map((feature) => (
            <div key={feature.title} className="rounded-lg border border-gray-200 p-5 dark:border-gray-800">
              <h3 className="font-semibold text-gray-900 dark:text-gray-100">{feature.title}</h3>
              <p className="mt-2 text-sm text-gray-600 dark:text-gray-400">{feature.body}</p>
            </div>
          ))}
        </div>
      </div>
    </section>
  );
}
