import StatCard from '../components/Dashboard/StatCard';

import {
  Calculator,
  CircleDollarSign,
  CircleQuestionMark,
  Mail,
  Users,
} from 'lucide-react';

export default function DashboardPage() {
  return (
    <main className="min-h-screen bg-gray-100 p-8">
      <section className="mx-auto max-w-6xl">
        <h1 className="text-3xl font-bold text-green-800">
          Dufferin Lawn Life Atlas
        </h1>

        <div className="mt-8 grid gap-6 md:grid-cols-4">
          <StatCard
            title="Real Green"
            icon={Users}
            href="https://zooey.serviceassistant.com/838579/Login"
          />
          <StatCard
            title="Quickbooks"
            icon={CircleDollarSign}
            href="https://qbo.intuit.com/app/homepage?locale=en-CA"
          />
          <StatCard
            title="EstimateTool"
            icon={Calculator}
            href="/Estimate/index.html"
          />
          <StatCard
            title="Email"
            icon={Mail}
            href="https://outlook.live.com/mail/0/"
          />
          <StatCard title="Coming Soon" icon={CircleQuestionMark} />
        </div>
      </section>
    </main>
  );
}
