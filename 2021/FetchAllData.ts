/**
 * Test
 */

interface Question {
  text: string; // 'Relasjonsdatabaser som Postgres, Oracle o.l.',
  topic: string; // topic: 'Relasjonsdatabaser',
  category: string; // category: 'Backend',
  id: string; // id: 'b1656d08-1f76-443e-a23f-b6179235da75' } },
}

interface Answer {
  knowledge: number; // knowledge: 3,
  motivation: number; //     motivation: 2,
  updatedAt: string; //     updatedAt: '2021-02-22T12:15:51.688Z',
  question: Question; //     question:
}

interface TaxonomyTree {
  [key: string]: string;
}

/**
 * Wrapper for the ugly UrlFetchApp.fetch function
 *
 * @param url string - the URL to fetch
 * @returns any - the object return from JSON.parse
 */
function _fetch(url: string): any {
  console.log('fetching url:', url);

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
  console.log(`status: ${status}. Continuing`);

  return JSON.parse(res.getContentText());
}

function transpose(a: any[][]): any[][] {
  return a[0].map((_, i) => a.map((x) => x[i]));
}

/**
 * Updates and writes data to the data sheet
 */
function generateDataSheet() {
  const sData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  const sNotAnswered = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('not answered');

  sData.clearContents();
  sNotAnswered.clearContents();

  const users = getUserList().map((u) => u[0]);
  const categories = getCategoriesData();

  let catMap = new Map();

  categories.forEach((e) => {
    catMap.set(e.id, e.text);
  });

  // console.log('catmap', catMap.keys());
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

  console.log('job questions:', jobQuestions);
  const questions = jobQuestions.concat(compQuestions);

  const all = getAllAnswersData()
    .map((u) => {
      let r = [u.email, u.username, u.updatedAt.slice(0, 10)];

      const answers = new Map();
      const seenJobs = new Set();
      let jobs = [];
      u.answers.forEach((a) => {
        if (typeof a.customScaleValue !== 'undefined') {
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

      compQuestions.forEach((q) => {
        const id = q[2];
        if (answers.has(id)) {
          const a = answers.get(id);
          r.push(a.knowledge);
        } else {
          r.push('');
        }
      });
      return r;
    })
    .sort((a, b) => (a[0] > b[0] ? 1 : -1));

  all.forEach((e) => {
    if (e.length > 156) {
      console.log('length:', e.length);
      console.log('length:', e);
    }
  });

  const answered = all.map((u) => u[0]);
  const notAnswered = users.filter((u) => !answered.includes(u)).map((u) => [u]);
  const transposed = transpose(questions);
  const updated = `Last updated: ${new Date().toLocaleString('se')}`;

  sData.getRange(3, 1, 1, 3).setValues([['email', 'user id', 'updated at']]);
  sData.getRange(1, 4, transposed.length, transposed[0].length).setValues(transposed);
  sData.getRange(4, 1, all.length, all[0].length).setValues(all);
  sData.getRange(1, 1).setValue(updated);

  sNotAnswered.getRange(3, 1, notAnswered.length, 1).setValues(notAnswered);
  sNotAnswered.getRange(1, 1).setValue(updated);
}

/**
 * Fetches complete list of users having completed the competency mapping survey
 *
 * @returns list of users in the competency mapping database
 * @customfunction
 */
function getUserList() {
  const res = UrlFetchApp.fetch(config.urls.users, {
    headers: {
      'x-api-key': config.apikey,
    },
  });

  const status = res.getResponseCode();
  if (status !== 200) {
    console.log(`status: ${status}. Aborting update.`);
    return;
  }
  console.log(`status: ${status}. Continuing`);

  const data = JSON.parse(res.getContentText());
  const output = data
    .map((user) => [user.attributes[0].Value, user.username])
    .filter((e) => e[0] !== 'user@user.user')
    .sort((a, b) => (a[0] > b[0] ? 1 : -1));

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
  const questions = _fetch(`${config.urls.catalogs}/${config.catalogs.latest}/questions`);
  const qlist = questions.map((q) => q.id).sort();

  // console.log(qlist);

  const answers = qlist.map((id) => {
    const found = data.answers.find((a) => id === a.question.id);
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
    a.question.category,
    a.knowledge,
    a.motivation,
  ]);

  return answers;
}

function getCategoriesData() {
  return _fetch(`${config.urls.catalogs}/${config.catalogs.latest}/categories`);
}

function getAllAnswersData() {
  return _fetch(config.urls.answers);
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

/**
 * Fetches latest question catalog. Currently hard coded to id of latest catalog
 *
 * @returns
 * @customfunction
 */
function getQuestions() {
  const data = _fetch(`${config.urls.catalogs}/${config.catalogs.latest}/questions`);
  const output = data
    .map((q) => [q.index, q.topic, q.text, q.type, q.id, q.categoryID])
    .sort((a, b) => (a[5] > b[5] ? 1 : -1));
  return output;
}
