import { createBrowserClient } from '@supabase/ssr';

export function createClient() {
  const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const supabaseKey =
    process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY ??
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

  if (!supabaseUrl) {
    throw new Error(
      'Missing NEXT_PUBLIC_SUPABASE_URL. Add it to Cloudflare Pages environment variables and redeploy.',
    );
  }

  if (!supabaseKey) {
    throw new Error(
      'Missing Supabase public key. Add NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY or NEXT_PUBLIC_SUPABASE_ANON_KEY to Cloudflare Pages environment variables and redeploy.',
    );
  }

  return createBrowserClient(supabaseUrl, supabaseKey);
}
