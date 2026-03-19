'use client';

import { useAuth } from '@/lib/useAuth';
import Header from '@/components/Header';

export default function DashboardPage() {
  const { session, loading, logout } = useAuth();

  if (loading || !session) return null;

  return (
    <div className="min-h-screen bg-gray-50">
      <Header session={session} onLogout={logout} />
      <main className="max-w-screen-xl mx-auto px-4 py-12 flex flex-col items-center justify-center gap-6">
        <div className="text-center">
          <h1 className="text-3xl font-bold text-gray-900 tracking-tight">Defy Field Execute</h1>
          <p className="text-gray-500 mt-2">Reporting platform — upload your Perigee exports to generate reports.</p>
        </div>
        <div className="bg-white border-2 border-dashed border-gray-200 rounded-2xl p-16 text-center text-gray-400 max-w-lg w-full">
          <p className="text-lg font-medium">Reports coming soon</p>
          <p className="text-sm mt-2">Upload functionality will be built once raw files are available.</p>
        </div>
      </main>
    </div>
  );
}
