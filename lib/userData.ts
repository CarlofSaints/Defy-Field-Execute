import fs from 'fs';
import path from 'path';

export interface User {
  id: string;
  name: string;
  email: string;
  password: string;
  isAdmin: boolean;
  forcePasswordChange: boolean;
  firstLoginAt: string | null;
  createdAt: string;
}

const FILE = path.join(process.cwd(), 'data', 'users.json');

export function loadUsers(): User[] {
  if (fs.existsSync(FILE)) {
    return JSON.parse(fs.readFileSync(FILE, 'utf-8'));
  }
  const env = process.env.DFE_USERS_JSON;
  if (env) return JSON.parse(env);
  return [];
}

export function saveUsers(users: User[]) {
  try {
    const dir = path.dirname(FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(FILE, JSON.stringify(users, null, 2));
  } catch {
    // Vercel serverless: writes don't persist — manage users via DFE_USERS_JSON env var
  }
}
