const fs = require('fs');
const path = require('path');
const { createLogger, format, transports } = require('winston');

const logsDir = path.resolve(__dirname, '..', 'logs');
fs.mkdirSync(logsDir, { recursive: true });

const fileFormat = format.combine(
  format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
  format.printf(({ timestamp, level, message }) => `[${timestamp}] ${level}: ${message}`)
);

const consoleFormat = format.combine(
  format.timestamp({ format: 'HH:mm:ss' }),
  format.colorize(),
  format.printf(({ timestamp, level, message }) => `[${timestamp}] ${level}: ${message}`)
);

const logger = createLogger({
  level: 'info',
  transports: [
    new transports.Console({ format: consoleFormat }),
    new transports.File({
      filename: path.join(logsDir, 'app.log'),
      format: fileFormat,
    }),
    new transports.File({
      filename: path.join(logsDir, 'error.log'),
      level: 'error',
      format: fileFormat,
    }),
  ],
});

module.exports = logger;
