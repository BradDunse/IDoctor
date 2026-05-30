const express = require('express');
const cors = require('cors');
const path = require('path');
const rewriteRoute = require('./api/rewrite');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50kb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/rewrite', rewriteRoute);

app.use(function(req, res) {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('The I Doctor running on port ' + PORT));
