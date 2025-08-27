const express = require('express');
const cors = require('cors');
const app = express();
const port = process.env.PORT || 3000;

// Allow CORS from the frontend dev server. Set ALLOWED_ORIGIN in env to customize.
const allowedOrigin = 'https://contabil.gfestatistica.com.br'
app.use(cors({ origin: allowedOrigin }))

app.use(express.json());

// Import routes
const razaoRouter = require('./routes/razao');
const dreRouter = require('./routes/dre');
const comparativoRouter = require('./routes/comparativo');
const balanceteRouter = require('./routes/balancete');
const companiesRouter = require('./routes/companies');
const statusRouter = require('./routes/status');

app.use('/razao', razaoRouter);
app.use('/dre', dreRouter);
app.use('/comparativo', comparativoRouter);
app.use('/balancete', balanceteRouter);
app.use('/companies', companiesRouter);
app.use('/status', statusRouter);

app.get('/', (req, res) => res.json({ status: 'ok' }));

app.listen(port, () => {
  console.log(`Server listening on port ${port} - CORS allowed for ${allowedOrigin}`);
});
