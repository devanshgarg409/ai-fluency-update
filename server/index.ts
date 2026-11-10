import 'dotenv/config';
import express, { Request, Response } from 'express';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import authRouter from './routes/auth.ts';
import teachersRouter from './routes/teachers.ts';
import studentsRouter from './routes/students.ts';
import schemaRouter from './routes/schema.ts';
import outlookRouter from './modules/outlook/outlook.routes.ts';
import { initializeDatabase } from './db/firebase.ts';

export const app = express();

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Initialize persistent DB & initial admin account
initializeDatabase().catch((err) => {
  console.error('Database initialization warning:', err);
});

// API Routes
app.use('/api/auth', authRouter);
app.use('/api/teachers', teachersRouter);
app.use('/api/students', studentsRouter);
app.use('/api/schema', schemaRouter);
app.use('/api/outlook', outlookRouter);

// Database re-sync endpoint to fetch all details directly from Firestore
app.post('/api/sync', async (req: Request, res: Response) => {
  try {
    const { syncDatabaseFromFirestore } = await import('./db/firebase.ts');
    const result = await syncDatabaseFromFirestore();
    return res.json(result);
  } catch (err: any) {
    return res.status(500).json({ error: 'Failed to resync database: ' + err.message });
  }
});

app.get('/api/sync', async (req: Request, res: Response) => {
  try {
    const { syncDatabaseFromFirestore } = await import('./db/firebase.ts');
    const result = await syncDatabaseFromFirestore();
    return res.json(result);
  } catch (err: any) {
    return res.status(500).json({ error: 'Failed to resync database: ' + err.message });
  }
});

// Health check endpoint
app.get('/api/health', (req: Request, res: Response) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    service: 'EduCore Academic Portal API',
  });
});

// Production static file serving (for Render / Vercel / standalone deployments)
const distPath = path.resolve(process.cwd(), 'dist');
if (fs.existsSync(distPath)) {
  app.use(express.static(distPath));
  app.get('*', (req: Request, res: Response) => {
    if (!req.url.startsWith('/api')) {
      res.sendFile(path.join(distPath, 'index.html'));
    }
  });
}

// Standalone server starter - only when executed as main entry
const isDirectRun =
  process.env.START_STANDALONE === 'true' ||
  (process.argv[1] &&
    (process.argv[1].includes('server/index') ||
      process.argv[1].includes('server.ts') ||
      process.argv[1].includes('server.js')));

if (isDirectRun) {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => {
    console.log(`EduCore Backend server listening on port ${PORT}`);
  });
}

export default app;
