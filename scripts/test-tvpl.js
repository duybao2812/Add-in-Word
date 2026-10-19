const https = require('https');

function testApi(lawId) {
  const url = `https://apimobile.thuvienphapluat.vn/Document/GetDocumentDetail?LawID=${lawId}&language=&token=`;
  return new Promise((resolve, reject) => {
    https.get(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X)',
        'Accept': 'application/json, text/plain, */*',
        'Origin': 'https://m.thuvienphapluat.vn',
        'Referer': 'https://m.thuvienphapluat.vn/'
      }
    }, res => {
      console.log('API Status:', res.statusCode);
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        const text = Buffer.concat(chunks).toString('utf8');
        try {
          const json = JSON.parse(text);
          resolve(json);
        } catch (e) {
          resolve({ raw: text.slice(0, 500) });
        }
      });
    }).on('error', reject);
  });
}

testApi(709388).then(data => {
  console.log('Keys in data:', Object.keys(data));
  if (data.Document) {
    console.log('Document Code:', data.Document.News_Code);
    console.log('Document Subject:', data.Document.News_Subject);
  }
  if (data.ContentVN) {
    console.log('ContentVN length:', data.ContentVN.length);
    console.log('ContentVN snippet:', data.ContentVN.slice(0, 400));
  } else {
    console.log('Full response:', JSON.stringify(data).slice(0, 600));
  }
}).catch(console.error);
