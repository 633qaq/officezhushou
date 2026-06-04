const config = require('./config');
const { app } = require('./app');
const { initDatabase, closeDatabase } = require('./database');
const logger = require('./utils/logger');

async function main() {
  try {
    await initDatabase();
    logger.info('Database initialized');

    const server = app.listen(config.port, () => {
      logger.info('========================================');
      logger.info('Office AI Assistant server is running');
      logger.info(`Port: ${config.port}`);
      logger.info(`Environment: ${config.nodeEnv}`);
      logger.info(`Web: http://localhost:${config.port}`);
      logger.info(`API: http://localhost:${config.port}/api`);
      logger.info('========================================');
    });

    const shutdown = (signal) => {
      logger.info(`Received ${signal}, shutting down`);
      server.close(() => {
        closeDatabase();
        logger.info('HTTP server closed');
        process.exit(0);
      });

      setTimeout(() => process.exit(1), 5000);
    };

    process.on('SIGINT', () => shutdown('SIGINT'));
    process.on('SIGTERM', () => shutdown('SIGTERM'));
  } catch (error) {
    logger.error('Server startup failed', error.message, error.stack);
    process.exit(1);
  }
}

main();
