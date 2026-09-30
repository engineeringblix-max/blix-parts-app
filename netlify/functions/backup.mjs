// Weekly backup e-mail: every Monday around 07:00 (Amsterdam summer time; 06:00 in winter).
import { runBackup, runOverdue } from './api.mjs';

export default async () => {
  const e = await runBackup(process.env, 'weekly');
  if (e) console.error('Weekly backup failed:', e); else console.log('Weekly backup sent');
  const o = await runOverdue(process.env, 'weekly');
  if (o) console.error('Overdue summary failed:', o); else console.log('Overdue summary done');
};

export const config = { schedule: '0 5 * * 1' };
