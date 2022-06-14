const EventEmitter = require('events');

class MyEmitter extends EventEmitter {}

const emiter = new MyEmitter();

emiter.on('scrapOn', () => {
  console.log('Scraping event triggered bc');
});

module.exports = {emiter}
