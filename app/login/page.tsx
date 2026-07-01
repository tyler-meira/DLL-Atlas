'use client';

import React, { useState } from 'react';
import { useRouter } from 'next/navigation';
import { createClient } from '@/lib/supabase/client';

export default function LoginPage() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const router = useRouter();

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    // Basic client-side validation
    if (!email || !password) {
      setError('Please fill in all fields.');
      setLoading(false);
      return;
    }

    try {
      const supabase = createClient();

      const { error: signInError } = await supabase.auth.signInWithPassword({
        email,
        password,
      });

      if (signInError) {
        setError(signInError.message);
        return;
      }

      router.push('/dashboard');
      router.refresh();
    } catch (err) {
      console.error('Login failed:', err);
      setError(err instanceof Error ? err.message : 'Something went wrong. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="flex min-h-screen items-center justify-center bg-[#17884b] px-4 py-12">
      <div className="w-full max-w-md space-y-8 rounded-lg bg-[#fbfdf9] p-8 shadow-[0_20px_60px_rgba(16,24,20,0.18)] ring-1 ring-[#dfe9e2]">
        <div>
          <h2 className="mt-6 text-center text-3xl font-bold tracking-tight text-[#101814]">
            Sign in to your account
          </h2>
        </div>

        <form className="mt-8 space-y-6" onSubmit={handleSubmit}>
          {error && (
            <div className="rounded-md border border-[#f4c6bd] bg-[#fff3f0] p-3 text-sm text-[#b42318]">
              {error}
            </div>
          )}

          <div className="space-y-4">
            <div>
              <label
                htmlFor="email-address"
                className="block text-sm font-medium text-[#101814]"
              >
                Email address
              </label>
              <input
                id="email-address"
                name="email"
                type="email"
                autoComplete="email"
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="mt-1 block w-full rounded-md border border-[#cfded4] bg-white px-3 py-2 text-[#101814] placeholder-[#7b8b81] shadow-sm focus:border-[#22a65a] focus:outline-none focus:ring-2 focus:ring-[#a8e0bd] sm:text-sm"
                placeholder="you@example.com"
              />
            </div>

            <div>
              <label
                htmlFor="password"
                className="block text-sm font-medium text-[#101814]"
              >
                Password
              </label>
              <input
                id="password"
                name="password"
                type="password"
                autoComplete="current-password"
                required
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                className="mt-1 block w-full rounded-md border border-[#cfded4] bg-white px-3 py-2 text-[#101814] placeholder-[#7b8b81] shadow-sm focus:border-[#22a65a] focus:outline-none focus:ring-2 focus:ring-[#a8e0bd] sm:text-sm"
                placeholder="Enter your password"
              />
            </div>
          </div>

          <div>
            <button
              type="submit"
              disabled={loading}
              className="group relative flex w-full justify-center rounded-md border border-transparent bg-[#166f3a] px-4 py-2 text-sm font-medium text-white shadow-sm hover:bg-[#105c30] focus:outline-none focus:ring-2 focus:ring-[#22a65a] focus:ring-offset-2 focus:ring-offset-[#fbfdf9] disabled:bg-[#7aa98a]"
            >
              {loading ? 'Signing in...' : 'Sign in'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
