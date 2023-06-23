// Date.prototype.getJulian = () => {
//   return Math.floor(this / 86400000 - this.getTimezoneOffset() / 1440 + 2440587.5);
// };

// Date.prototype.oaDate = () => {
//   return (this - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
// };

const getJulian = (date) => {
  return Math.floor(date / 86400000 - date.getTimezoneOffset() / 1440 + 2440587.5);
};

const oaDate = (date) => {
  return (date - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
};

export { getJulian, oaDate };
