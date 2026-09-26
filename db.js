const mysql = require('mysql2/promise');

const pool = mysql.createPool({
  host: process.env.DB_HOST,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME,
  port: Number(process.env.DB_PORT) || 4000,
  ssl: {
    rejectUnauthorized: true   // <-- இது TiDB Cloud-க்கு மிக முக்கியம்!
  },
  waitForConnections: true,
  connectionLimit: 10,
  connectTimeout: 60000,
  enableKeepAlive: true,
  keepAliveInitialDelay: 10000
});

// Test Connection with full error logging
pool.getConnection()
  .then(conn => {
    console.log("✅ TiDB Connected successfully!");
    conn.release();
  })
  .catch(err => {
    console.error("MySQL Connection Error Full Details:", err);
  });

module.exports = pool;