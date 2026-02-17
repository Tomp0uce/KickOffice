/**
 * Backend implementation of officeRichText
 */
function formatTextForOffice(text) {
  if (!text) return [];
  return text.split('\n').map(line => line.trim()).filter(line => line.length > 0);
}

function isHtml(text) {
  return /<[a-z][\s\S]*>/i.test(text);
}

module.exports = {
  formatTextForOffice,
  isHtml
};
