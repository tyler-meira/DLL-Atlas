import StatCard from "./components/Dashboard/StatCard";

import {
  Users,
  Calculator
} from "lucide-react";

export default function Home() {
  return (
    <main className="min-h-screen bg-gray-100 p-8">
      <section className="mx-auto max-w-6xl">
        <h1 className="text-3xl font-bold text-green-800">
          Dufferin Lawn Life Dashboard
        </h1>

        <div className="mt-8 grid gap-6 md:grid-cols-4">
          <StatCard title="Customers" icon={Users} />
          <StatCard title="EstimateTool" icon={Calculator} href="Estimate/index.html" />
        </div>
      </section>
    </main>
  );
}