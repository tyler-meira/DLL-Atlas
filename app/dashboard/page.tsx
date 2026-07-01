'use client';

import { useEffect, useState } from 'react';
import { useRouter } from 'next/navigation';
import StatCard from '../components/Dashboard/StatCard';
import { createClient } from '@/lib/supabase/client';

import {
  Calculator,
  CircleDollarSign,
  CircleQuestionMark,
  Mail,
  Users,
} from 'lucide-react';

export default function DashboardPage() {
  const router = useRouter();
  const [isAllowed, setIsAllowed] = useState(false);

  useEffect(() => {
    let isMounted = true;
    const supabase = createClient();

    supabase.auth
      .getSession()
      .then(({ data }) => {
        if (!isMounted) {
          return;
        }

        if (!data.session) {
          router.replace('/login');
          return;
        }

        setIsAllowed(true);
      })
      .catch(() => {
        if (isMounted) {
          router.replace('/login');
        }
      });

    return () => {
      isMounted = false;
    };
  }, [router]);

  if (!isAllowed) {
    return (
      <main className="flex min-h-screen items-center justify-center bg-[#17884b] px-4">
        <p className="text-sm font-medium text-white">Checking access...</p>
      </main>
    );
  }

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
