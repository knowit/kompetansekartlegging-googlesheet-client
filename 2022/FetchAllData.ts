/**
 * Wrapper for the ugly UrlFetchApp.fetch function
 *
 * @param url string - the URL to fetch
 * @returns any - the object return from JSON.parse
 */
function _fetch(url: string): any {
  const res = UrlFetchApp.fetch(url, {
    headers: {
      'x-api-key': config.apikey,
    },
  });
  const status = res.getResponseCode();
  if (status !== 200) {
    console.log(`status: ${status}. Aborting update.`);
    return;
  }

  return JSON.parse(res.getContentText());
}

function transpose(a: any[][]): any[][] {
  return a[0].map((_, i) => a.map((x) => x[i]));
}

function getUserBlocklist(): string[] {
  const sBlocklist = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('user blocklist');
  if (sBlocklist === null) throw new TypeError('Spreadsheet sheet user blocklist is null');

  return sBlocklist.getDataRange().getValues().flat().slice(2);
}

/**
 * Updates and writes data to the data sheet
 */
function generateDataSheet() {
  const sData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  if (sData === null) throw new TypeError('Spreadsheet sheet data is null');

  const sNotAnswered = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('not answered');
  if (sNotAnswered === null) throw new TypeError('Spreadsheet sheet not answered is null');

  sData.clearContents();
  sNotAnswered.clearContents();

  const users = getUserList();
  const categories: Category[] = getCategoriesData();

  let catMap = new Map();

  categories.forEach((e: Category) => {
    catMap.set(e.id, e.text);
  });

  const allQuestions = getQuestions();
  const compQuestions = allQuestions
    .filter((a) => a[3] === 'knowledgeMotivation')
    .sort((a, b) => a[0] - b[0])
    .map((e) => {
      if (catMap.has(e[5])) {
        e.push(catMap.get(e[5]));
      }
      return [e[6], e[1], e[4]];
    });

  const jobQuestions = allQuestions
    .filter((a) => a[3] === 'customScaleLabels')
    .sort((a, b) => (a[4] > b[4] ? 1 : -1))
    .map((e) => {
      if (catMap.has(e[5])) e.push(catMap.get(e[5]));
      return [e[6], e[1], e[4]];
    });

  const questions = jobQuestions.concat(compQuestions, compQuestions);
  const blocklist = getUserBlocklist();
  const all = getAllAnswersData()
    .map((u: UserAnswers) => {
      let r = [u.email, u.username, u.updatedAt.slice(0, 10)];

      const answers = new Map();
      const seenJobs = new Set();

      let jobs: any[] = [];
      u.answers.filter((a) => a.hasOwnProperty("question")).forEach((a) => {
        if (typeof a.customScaleValue !== 'undefined') {
          // workaround for a bug in backend where some customScaleValues are
          // duplicates for some users
          if (!seenJobs.has(a.question.id)) {
            jobs.push([a.question.id, a.customScaleValue]);
            seenJobs.add(a.question.id);
          }
        } else {
          answers.set(a.question.id, {
            knowledge: a.knowledge,
            motivation: a.motivation,
          });
        }
      });

      jobs = jobs.sort((a, b) => (a[0] > b[0] ? 1 : -1)).map((a) => a[1]);
      while (jobs.length < 2) {
        jobs.push('');
      }
      r.push(...jobs);

      // Add all the knowledge values first
      compQuestions.forEach((q) => {
        const id = q[2];
        if (answers.has(id)) {
          const a = answers.get(id);
          r.push(a.knowledge);
        } else {
          r.push('');
        }
      });

      // Add all the motivation values
      compQuestions.forEach((q) => {
        const id = q[2];
        if (answers.has(id)) {
          const a = answers.get(id);
          r.push(a.motivation);
        } else {
          r.push('');
        }
      });
      return r;
    })
    .sort((a, b) => (a[0] > b[0] ? 1 : -1)) // Sort by email
    .filter((u) => !blocklist.includes(u[0])); // Remove users who have quit

  // figure out who has not answered
  const answered = all.map((u) => u[0]);
  const notAnswered = users.filter((u) => !answered.includes(u)).map((u) => [u]);

  // transpose questions to print horizontally
  const transposed = transpose(questions);

  sData.getRange(4, 1, 1, 3).setValues([['email', 'user id', 'updated at']]);
  sData.getRange(2, 4, transposed.length, transposed[0].length).setValues(transposed);
  sData.getRange(5, 1, all.length, all[0].length).setValues(all);
  sData.getRange(1, 6).setValue('Kompetanse');
  sData.getRange(1, 157).setValue('Motivasjon');

  // set last updated string
  const updated = `Last updated: ${new Date().toLocaleString('se')}`;
  sData.getRange(1, 1).setValue(updated);

  if (notAnswered.length > 0) {
    sNotAnswered.getRange(3, 1, notAnswered.length, 1).setValues(notAnswered);
    sNotAnswered.getRange(2, 1).setValue('email');
    sNotAnswered.getRange(1, 1).setValue(updated);
  }
}

/**
 * Fetches list of user emails last synchronised from AD, sorted alphabetically.
 *
 * @returns any[]
 */
function getUserList(): any[] {
  const data = _fetch(config.urls.users);
  const output = data
    .map((user: User) => <string[]>[user.attributes[0].Value, user.username])
    .map((e: string[]) => e[0])
    .filter((e: string) => e !== 'user@user.user')
    .sort();

  return output;
}

type KnowledgeMotivation = 'knowledge' | 'motivation';

/**
 * Fetches the latest answers for user by id.
 *
 * @param username string
 * @returns
 * @customfunction
 */
function getAnswersForUsername(username: string, type: KnowledgeMotivation) {
  const data = _fetch(`${config.urls.answers}/${username}/newest`);
  const questions: Question[] = _fetch(`${config.urls.catalogs}/${config.catalogs.latest}/questions`);
  const qlist = questions.map((q) => q.id).sort();

  const answers = qlist.map((id) => {
    const found = data.answers.find((a: Answer) => id === a.question.id);
    if (!found) return '';
    if (type === 'knowledge') {
      return found.knowledge ? found.knowledge : '';
    }
    if (type === 'motivation') {
      return found.motivation ? found.motivation : '';
    }
    return '';
  });
  const output = [data.updatedAt].concat(answers);

  return output;
}

/**
 * Fetches all the answers for a user.
 *
 * @param username string
 * @returns array range of answers unordered
 * @customfunction
 */
function getAllAnswersForUsername(username: string): any {
  const data = _fetch(`${config.urls.answers}/${username}/newest`);

  const answers = data.answers.map((a: Answer) => [
    a.question.id,
    a.updatedAt,
    a.question.topic,
    a.question.categoryID,
    a.knowledge,
    a.motivation,
  ]);

  return answers;
}

/**
 * Fetches and returns the list of categories sorted according to index.
 * 
 */
function getCategoriesData(): Category[] {
  const data = _fetch(`${config.urls.catalogs}/${config.catalogs.latest}/categories`);
  return data.sort((a: Category, b: Category) => a.index - b.index);
}

function getAllAnswersData(): UserAnswers[] {
  const data = _fetch(config.urls.answers);
  return data;
}

/**
 * Fetches latest categories. Currently hard coded to id of latest catalog
 *
 * @returns
 * @customfunction
 */
function getCategories() {
  const output = getCategoriesData()
    .map((c) => [c.index, c.text, c.id, c.description])
    .sort((a, b) => (a[0] > b[0] ? 1 : -1));

  return output;
}

function getQuestionsData(): Question[] {
  return _fetch(`${config.urls.catalogs}/${config.catalogs.latest}/questions`);
}

/**
 * Fetches latest question catalog. Currently hard coded to id of latest catalog.
 *
 * @returns
 * @customfunction
 */
function getQuestions(): any[] {
  const data: Question[] = getQuestionsData();
  const output = data
    .map((q) => [q.index, q.topic, q.text, q.type, q.id, q.categoryID])
    .sort((a, b) => (a[5] > b[5] ? 1 : -1));
  return output;
}
