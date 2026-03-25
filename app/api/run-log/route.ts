import { NextResponse } from 'next/server';
import { loadRunLog } from '@/lib/runLogData';

export async function GET() {
  return NextResponse.json(loadRunLog());
}
