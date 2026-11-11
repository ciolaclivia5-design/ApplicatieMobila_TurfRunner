import { Pool } from 'pg';

const RENDER_DATABASE_URL = "postgresql://livia05_user:5fqlvgOxr7EhN2kzEsrDG1kRTYOxbbaV@dpg-db73lmt9fdbs73etvkeg-a.frankfurt-postgres.render.com/livia05"

const pool = new Pool({
  connectionString: RENDER_DATABASE_URL,
  ssl: {
      rejectUnauthorized: false
  }
});

export default pool;