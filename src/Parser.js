import fs from 'fs/promises';
import XLSX from 'xlsx';
import path from 'path';
import { MongoClient } from 'mongodb';

export default class Parser {
  settings = {
    daysOfWeek: 'A',
    group: 'A14',
    start: 16,
    length: 13,
    educationDays: 6,
    odd: {
      lessonNumber: 'B',
      time: 'C',
      classroom: 'D',
      lessonType: 'E',
      teacher: 'F',
      lesson: 'G',
    },
    even: {
      lessonNumber: 'L',
      time: 'M',
      classroom: 'K',
      lessonType: 'J',
      teacher: 'I',
      lesson: 'H',
    },
  };

  #client;
  #db;
  #schedules;
  #groups;
  #teachers;
  #lessons;

  constructor() {
    this.#client = new MongoClient('mongodb://localhost:27017', {
      maxIdleTimeMS: 1000,
      keepAlive: false,
    });
    this.#db = this.#client.db('kosygin_schedule');
    this.#schedules = this.#db.collection('schedules');
    this.#teachers = this.#db.collection('teachers');
    this.#lessons = this.#db.collection('lessons');
    this.#groups = this.#db.collection('groups');
  }

  async parseAll() {
    await this.#lessons.deleteMany({});
    await this.#schedules.deleteMany({});
    await this.#teachers.deleteMany({});
    await this.#groups.deleteMany({});

    let groupCount = 0;

    const rawFileData = await fs.readFile(path.resolve('..', 'filedata.json'));
    const fileData = JSON.parse(rawFileData);
    const groups = this.#db.collection('groups');
    fileData.forEach((file) => {
      const table = XLSX.readFile(file.path);
      table.SheetNames.forEach((group) => {
        if (
          group.toLowerCase().includes('майнор') ||
          group.toLowerCase().includes('уг_')
        ) {
          return;
        }
        (async () => {
          await groups.updateOne(
            { group },
            {
              $set: {
                instituteId: file.instituteId,
                instituteName: file.instituteName,
                course: file.course,
              },
            },
            { upsert: true }
          );

          await this.#parseOne(table.Sheets[group], group, file.instituteId);
        })();
      });
    });
    return this.#client;
  }

  #makeCouple(arr) {
    const size = 2;
    const result = [];
    for (let i = 0; i < Math.ceil(arr.length / size); i++) {
      result[i] = arr.slice(i * size, i * size + size);
    }
    return result;
  }

  #getDayOfWeek(dayOfWeek) {
    switch (dayOfWeek) {
      case 'ПН':
        return 'Понедельник';
      case 'ВТ':
        return 'Вторник';
      case 'СР':
        return 'Среда';
      case 'ЧТ':
        return 'Четверг';
      case 'ПТ':
        return 'Пятница';
      case 'СБ':
        return 'Суббота';
      default:
        return dayOfWeek;
    }
  }

  async #parseOne(page, group, instituteId) {
    console.log(group);
    let currentLine = 10;
    while (
      !page['A' + currentLine]?.v?.trim().toLowerCase().includes('понедельник')
    ) {
      currentLine++;
    }
    currentLine++;
    const schedule = [];

    for (let i = 0; i < this.settings.educationDays; i++) {
      const dayOfWeekSchedule = {
        location: page['A' + (currentLine - 1)]?.v.trim(),
        dayOfWeek: this.#getDayOfWeek(page['A' + currentLine]?.v?.trim()),
        odd: [],
        even: [],
      };

      let lessonNumber = 1;
      do {
        dayOfWeekSchedule.even.push({
          lessonNumber,
          time: page[this.settings.even.time + currentLine]?.v,
          classroom: page[this.settings.even.classroom + currentLine]?.v,
          lessonType: page[this.settings.even.lessonType + currentLine]?.v,
          teacher: page[this.settings.even.teacher + currentLine]?.v,
          lesson: page[this.settings.even.lesson + currentLine]?.v,
        });
        dayOfWeekSchedule.odd.push({
          lessonNumber,
          time: page[this.settings.odd.time + currentLine]?.v,
          classroom: page[this.settings.odd.classroom + currentLine]?.v,
          lessonType: page[this.settings.odd.lessonType + currentLine]?.v,
          teacher: page[this.settings.odd.teacher + currentLine]?.v,
          lesson: page[this.settings.odd.lesson + currentLine]?.v,
        });

        if (page[`${this.settings.even.lesson}${currentLine}`]?.v) {
          await this.#lessons.insertOne({
            group,
            isEven: true,
            dayOfWeek: i,
            lessonNumber,
            time: page[this.settings.even.time + currentLine]?.v,
            classroom: page[this.settings.even.classroom + currentLine]?.v,
            lessonType: page[this.settings.even.lessonType + currentLine]?.v,
            teacher: page[this.settings.even.teacher + currentLine]?.v,
            lesson: page[this.settings.even.lesson + currentLine]?.v,
          });
        }
        if (page[`${this.settings.odd.lesson}${currentLine}`]?.v) {
          await this.#lessons.insertOne({
            group,
            dayOfWeek: i,
            isEven: false,
            lessonNumber,
            time: page[this.settings.odd.time + currentLine]?.v,
            classroom: page[this.settings.odd.classroom + currentLine]?.v,
            lessonType: page[this.settings.odd.lessonType + currentLine]?.v,
            teacher: page[this.settings.odd.teacher + currentLine]?.v,
            lesson: page[this.settings.odd.lesson + currentLine]?.v,
          });
        }

        currentLine++;
        lessonNumber++;
      } while (page['B' + currentLine]);
      currentLine++;
      dayOfWeekSchedule.even = this.#makeCouple(dayOfWeekSchedule.even);
      dayOfWeekSchedule.odd = this.#makeCouple(dayOfWeekSchedule.odd);
      schedule.push(dayOfWeekSchedule);
    }
    this.#schedules.updateOne(
      { group },
      {
        $set: { schedule },
      },
      { upsert: true }
    );
    return schedule;
  }

  async #saveJSON(_path, filename, data) {
    try {
      await fs.access(_path);
    } catch (e) {
      if (e.code === 'ENOENT') {
        await fs.mkdir(_path, { recursive: true });
      }
    }
    await fs.writeFile(path.resolve(_path, filename), JSON.stringify(data));
  }
}
