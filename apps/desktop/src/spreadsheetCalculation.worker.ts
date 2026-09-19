self.onmessage = () => {
  self.postMessage({ error: 'Local workbook calculation engine is not available.' });
};
