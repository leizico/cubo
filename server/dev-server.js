import { serve } from '@hono/node-server';
import app from './app.js';

const port = Number(process.env.API_PORT || 3001);

serve({ fetch: app.fetch, port }, () => {
  console.log(`BI Cubo API em http://localhost:${port}`);
});
