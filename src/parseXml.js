const fs = require('fs');
const { XMLParser } = require('fast-xml-parser');

function parseXmlFile(xmlPath) {
  const xml = fs.readFileSync(xmlPath, 'utf8');
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_'
  });
  const json = parser.parse(xml);
  return json;
}

module.exports = { parseXmlFile };
