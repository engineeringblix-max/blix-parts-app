// Weekly backup e-mail: every Monday around 07:00 (Amsterdam summer time; 06:00 in winter).
import { runBackup } from './api.mjs';

export default async () => {
  const e = await runBackup(process.env, 'weekly');
  if (e) console.error('Weekly backup failed:', e); else console.log('Weekly backup sent');
};

export const config = { schedule: '0 5 * * 1' };
