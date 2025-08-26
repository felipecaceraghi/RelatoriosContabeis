const express = require('express');
const app = express();
const port = process.env.PORT || 3000;

app.use(express.json());

// Import routes
const razaoRouter = require('./routes/razao');
const dreRouter = require('./routes/dre');
const comparativoRouter = require('./routes/comparativo');
const balanceteRouter = require('./routes/balancete');

app.use('/razao', razaoRouter);
app.use('/dre', dreRouter);
app.use('/comparativo', comparativoRouter);
app.use('/balancete', balanceteRouter);

app.get('/', (req, res) => res.json({ status: 'ok' }));

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
