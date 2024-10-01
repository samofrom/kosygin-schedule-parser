import { parse } from 'node-html-parser';
import puppeteer, { errors } from 'puppeteer';
import fs from 'fs/promises';
import { createWriteStream } from 'fs';
import axios from 'axios';
import path from 'path';
import stream from 'stream/promises';

export default class Scraper {
  /* Institute ids
    10 - Колледж
    3 - Институт дизайна
    4 - Институт искусств
    16 - ИНСТИТУТ МЕХАТРОНИКИ И РОБОТОТЕХНИКИ
    17 - ИНСТИТУТ ИНФОРМАЦИОННЫХ ТЕХНОЛОГИЙ И ЦИФРОВОЙ ТРАНСФОРМАЦИИ
    7 - ИНСТИТУТ СОЦИАЛЬНОЙ ИНЖЕНЕРИИ
    8 - ИНСТИТУТ ХИМИЧЕСКИХ ТЕХНОЛОГИЙ И ПРОМЫШЛЕННОЙ ЭКОЛОГИИ
    9 - ИНСТИТУТ ЭКОНОМИКИ И МЕНЕДЖМЕНТА
    18 - ТЕХНОЛОГИЧЕСКИЙ ИНСТИТУТ ТЕКСТИЛЬНОЙ И ЛЕГКОЙ ПРОМЫШЛЕННОСТИ
    6 - ИНСТИТУТ СЛАВЯНСКОЙ КУЛЬТУРЫ
    2 - ИНСТИТУТ «АКАДЕМИЯ ИМЕНИ МАЙМОНИДА»
    14 - ИНСТИТУТ МЕЖДУНАРОДНОГО ОБРАЗОВАНИЯ
    11 - МАГИСТРАТУРА
  */
  #settings = {
    excludeInstitutes: [2, 3, 6, 11, 14, 9], //7 9 10 16 17 18
    fixEducationForms: [4, 8, 9],
    excludeForms: ['очно-заочная', 'заочная'],
  };
  #html;
  #links;

  async initialize() {
    this.#html = await this.#getHTML();
    this.#links = await this.#getLinks(this.#html);
    await this.downloadAll();
  }

  async #getHTML() {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(
      'https://kosygin-rgu.ru/1AppRGU/rguschedule/Embed/embedClassSchedule.aspx#formIndexMaster',
      { waitUntil: 'networkidle0' }
    );
    const data = await page.evaluate(
      () => document.querySelector('*').outerHTML
    );
    await browser.close();
    return data;
  }

  async #getLinks() {
    console.log('Getting links...');
    const html = parse(this.#html);
    const classSchedule = html.getElementById('ClassSchedule'); //Root div
    const idUnitClasses = classSchedule.querySelectorAll(
      'div[id*="idUnitClass-"]'
    ); //div containing the name of the institute
    const units = idUnitClasses.map((idUnitClass) => {
      const id = idUnitClass.id.split('-')[1]; //Institute id

      if (this.#settings.excludeInstitutes.includes(Number(id))) return; //skip institutes

      const cntUnitClass = classSchedule.querySelector(
        `div[id*="cntUnitClass-${id}"]`
      ); //div containing forms of education, courses and links to schedule

      const instituteName = idUnitClass.querySelector('label').textContent;

      const panelGroups = cntUnitClass.querySelectorAll('.panel-group');

      const forms = panelGroups.map((panelGroup) => {
        let formName = panelGroup.querySelector('strong');
        if (!formName) return; //skip empty panelgroups

        formName = formName.innerText.split(' ')[2];
        formName = formName.charAt(0) + formName.slice(1).toLowerCase();
        if (this.#settings.excludeForms.includes(formName.toLowerCase()))
          return;
        const tables = this.#settings.fixEducationForms.includes(+id) //fix for college
          ? panelGroup.querySelectorAll('.panel-body table').slice(2)
          : //exclude educational process schedule
            panelGroup.querySelectorAll('.panel-body table').slice(1);
        const courses = tables.map((table) => {
          console.log(instituteName);
          const td = table.querySelectorAll('td').slice(-1)[0];
          const a = td.querySelector('a');
          const course = a.textContent.toLowerCase();
          const link = a.getAttribute('href');
          const lastChange = td
            .querySelector('span')
            .textContent.split(' ')
            .slice(-1)[0];
          return {
            course,
            lastChange,
            link,
          };
        });
        return {
          [formName]: courses,
        };
      });
      return {
        id,
        instituteName:
          instituteName.charAt(0) + instituteName.slice(1).toLowerCase(),
        forms: forms.filter((form) => !!form), //fix empty form
      };
    });
    console.log('Links received');
    return units.filter((unit) => !!unit); //fix excluded units
  }

  async downloadAll() {
    console.log('Downloading spreadsheets...');
    const fileData = [];
    this.#links.forEach((unit) => {
      const institutePath = unit.instituteName
        .split(' ')
        .join('-')
        .replaceAll(';', '');
      unit.forms.forEach((form) => {
        Object.keys(form).forEach((formName) => {
          form[formName].forEach(({ course, lastChange, link }) => {
            const coursePath = course.split(' ').join('_');
            (async () => {
              fileData.push({
                instituteId: unit.id,
                instituteName: unit.instituteName,
                form: formName,
                lastChange,
                course,
                path: path.resolve(
                  '..',
                  'xlsx',
                  institutePath,
                  formName,
                  `${coursePath}.xlsx`
                ),
              });
              try {
                await fs.mkdir(
                  path.resolve('..', 'xlsx', institutePath, formName),
                  {
                    recursive: true,
                  }
                );
                const writer = createWriteStream(
                  path.resolve(
                    '..',
                    'xlsx',
                    institutePath,
                    formName,
                    `${coursePath}.xlsx`
                  )
                );
                const response = await axios.get(link, {
                  responseType: 'stream',
                });
                await stream.pipeline(response.data.pipe(writer));
                writer.on('finish', () => writer.end());
              } catch (e) {
                if (e.code === 'ENOENT') {
                  await fs.mkdir(
                    path.resolve('..', 'xlsx', institutePath, formName)
                  );
                  console.log(e);
                }
              }
            })();
          });
        });
      });
    });
    try {
      await fs.writeFile(
        path.resolve('..', 'filedata.json'),
        JSON.stringify(fileData)
      );
    } catch (e) {
      console.log(e);
    }
    console.log('Spreadsheets downloaded successfully');
    return fileData;
  }
}
