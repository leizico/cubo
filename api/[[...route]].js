import { handle } from 'hono/vercel';
import app from '../server/app.js';

export const config = {
  runtime: 'nodejs',
  maxDuration: 30,
};

export default handle(app);
