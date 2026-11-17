import { Pool } from 'pg';

const RENDER_DATABASE_URL = "postgresql://livia:al7nizGGjT9IUGwwQPq3P4iKt6YzP200@dpg-dbtvdt56e07s73bbln30-a.frankfurt-postgres.render.com/livia_database"

const pool = new Pool({
  connectionString: RENDER_DATABASE_URL,
  ssl: {
      rejectUnauthorized: true
  }
});

export default pool;