'use client';

import { Session } from '@/lib/useAuth';
import Image from 'next/image';
import Link from 'next/link';
import { useState } from 'react';

interface HeaderProps {
  session: Session;
  onLogout: () => void;
}

export default function Header({ session, onLogout }: HeaderProps) {
  const [defyErr, setDefyErr] = useState(false);
  const [atomicErr, setAtomicErr] = useState(false);

  return (
    <>
      <header className="bg-white border-b border-gray-200 shadow-sm sticky top-0 z-30">
        <div className="max-w-screen-xl mx-auto px-4 h-16 flex items-center justify-between gap-4">
          {/* Left: Defy logo + title */}
          <div className="flex items-center gap-3 min-w-0">
            {!defyErr ? (
              <Image
                src="/defy-logo.png"
                alt="Defy"
                width={80}
                height={32}
                className="object-contain"
                onError={() => setDefyErr(true)}
              />
            ) : (
              <span className="font-bold text-[#E31837] text-lg tracking-widest">DEFY</span>
            )}
            <div className="hidden sm:block border-l border-gray-200 pl-3">
              <p className="font-bold text-gray-900 text-sm leading-tight">Field Execute</p>
              <p className="text-xs text-gray-400 leading-tight">Reporting Platform</p>
            </div>
          </div>

          {/* Right: Atomic logo + user + controls */}
          <div className="flex items-center gap-3 shrink-0">
            {!atomicErr ? (
              <Image
                src="/atomic-logo.png"
                alt="Atomic Marketing"
                width={100}
                height={28}
                className="object-contain hidden sm:block"
                onError={() => setAtomicErr(true)}
              />
            ) : (
              <span className="text-xs font-semibold text-gray-500 hidden sm:block">ATOMIC MARKETING</span>
            )}

            <div className="w-px h-6 bg-gray-200 hidden sm:block" />

            <div className="hidden sm:block text-right">
              <p className="text-sm font-medium text-gray-800 leading-tight">{session.name}</p>
              <p className="text-xs text-gray-400 leading-tight">{session.email}</p>
            </div>

            {session.isAdmin && (
              <Link
                href="/admin/users"
                className="text-xs bg-gray-100 hover:bg-gray-200 text-gray-700 px-3 py-1.5 rounded font-medium transition-colors"
              >
                Users
              </Link>
            )}

            <button
              onClick={onLogout}
              className="text-xs bg-[#E31837] hover:bg-[#c01430] text-white px-3 py-1.5 rounded font-medium transition-colors"
            >
              Sign Out
            </button>
          </div>
        </div>
      </header>

      {/* Perigee fixed overlay bottom-right */}
      <div className="fixed bottom-4 right-4 z-40 opacity-60 hover:opacity-100 transition-opacity">
        <Image
          src="/perigee-logo.jpg"
          alt="Perigee"
          width={80}
          height={24}
          className="object-contain"
          onError={() => {}}
        />
      </div>
    </>
  );
}
