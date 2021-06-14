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

  console.log(qlist);

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

/**
 * Fetches latest categories. Currently hard coded to id of latest catalog
 *
 * @returns
 * @customfunction
 */
function getCategories() {
  const data = _fetch(`${config.urls.catalogs}/${config.catalogs.latest}/categories`);
  const output = data.map((c) => [c.index, c.text, c.id, c.description]).sort((a, b) => (a[0] > b[0] ? 1 : -1));

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
