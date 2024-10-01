import Scraper from './Scraper.js';
import Parser from './Parser.js';
import path from 'path';

const scraper = new Scraper();

const parser = new Parser();

await scraper.initialize();
(async () => {
  await parser.parseAll();
})();
