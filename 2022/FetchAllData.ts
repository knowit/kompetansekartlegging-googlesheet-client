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

  // get categories ordered by index (order in web app)
  const categories: Category[] = getCategoriesData();

  // create a lookup map to match ID's with data
  // let catMap = new Map();
  // categories.forEach((e: Category) => {
  //   catMap.set(e.id, e.text);
  // });

  // All questions
  const allQuestions = getQuestions();

  // Questions directly releated to competency: knowledge + motivation
  const compQuestions = allQuestions
    .filter((a) => a[3] === 'knowledgeMotivation')
    .sort((a, b) => a[0] - b[0]) // sort by question index
    .map((e) => {
      const category = categories.find((c) => c.id === e[5]);
      if (typeof category !== "undefined") {
        e.push(category.text, category);
        return [e[6], e[1], e[4], e[7]];
      }
    })
    .sort((a: any, b: any) => a[3].index - b[3].index); // sort by category index. FIX THIS.

  // Questions related to job funcions: job rotation, softskills
  const jobQuestions = allQuestions
    .filter((a) => a[3] === 'customScaleLabels')
    // sort by question id:
    .sort((a, b) => (a[4] > b[4] ? 1 : -1))
    // map replace category_id with category name:
    .map((e) => {
      const category = categories.find((c) => c.id === e[5])
      if (typeof category !== "undefined") {
        e.push(category.text, category);
      }
      return [e[6], e[1], e[4], e[7]];
    })
    // sort by category name:
    .sort((a, b) => a[0] > b[0] ? 1 : -1);

  const questions = jobQuestions.concat(compQuestions, compQuestions);
  const blocklist = getUserBlocklist();
  const allUsersAnswers = getAllUserQuestionAnswers();

  const answerMatrix = allUsersAnswers.map((u: UserQuestionAnswers) => {
    let r = [u.email, u.username, u.updatedAt.slice(0, 10)];

    const answers = new Map();
    const jobs = new Map();

    u.answers.forEach((a) => {
      if (typeof a.customScaleValue !== 'undefined') {
        // maybe still a bug: workaround for a bug in backend where some customScaleValues are
        // duplicates for some users
        jobs.set(a.questionId, a);
      } else {
        answers.set(a.questionId, a);
      }
    });

    for (const q of jobQuestions) {
      if (typeof q === "undefined") {
        throw new TypeError("Q is undefined");
      }
      // add all the job questions
      if (jobs.has(q[2])) {
        const j = jobs.get(q[2]);
        r.push(j.customScaleValue)
      } else {
        r.push("")
      }
    }

    for (const q of compQuestions) {
      // Add all the knowledge values first
      if (typeof q === "undefined") {
        throw new TypeError("Q is undefined");
      }
      if (answers.has(q[2])) {
        const a = answers.get(q[2]);
        r.push(a.knowledge);
      } else {
        r.push('');
      }
    }

    for (const q of compQuestions) {
      // Add all the motivation values
      if (typeof q === "undefined") {
        throw new TypeError("Q is undefined");
      }

      const id = q[2];
      if (answers.has(id)) {
        const a = answers.get(id);
        r.push(a.motivation);
      } else {
        r.push('');
      }
    }

    return r;
  })
    .sort((a, b) => (a[0] > b[0] ? 1 : -1)) // Sort by email
    .filter((u) => !blocklist.includes(u[0])); // Remove users who have quit

  // figure out who has not answered
  const answered = answerMatrix.map((u) => u[0]);
  const notAnswered = users.filter((u) => !answered.includes(u)).map((u) => [u]);

  // transpose questions to print horizontally
  const transposed = transpose(questions);

  sData.getRange(4, 1, 1, 3).setValues([['email', 'user id', 'updated at']]);
  sData.getRange(2, 4, transposed.length, transposed[0].length).setValues(transposed);
  sData.getRange(5, 1, answerMatrix.length, answerMatrix[0].length).setValues(answerMatrix);
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
 * Fetches and returns the list of categories sorted according to index.
 * 
 */
function getCategoriesData(): Category[] {
  const data = _fetch(`${config.urls.catalogs}/${config.catalogs.latest}/categories`);
  return data.sort((a: Category, b: Category) => (a.index === b.index) ? (a.text < b.text) : (a.index - b.index));
}

/**
 * Fetches answers for all users. 
 * Returns list sorted by email
 * 
 * @returns UserAnswers[]
 */
function getAllAnswersData(): UserAnswers[] {
  const data = _fetch(config.urls.answers).sort((a: UserAnswers, b: UserAnswers) => a.email > b.email ? 1 : -1);
  return data;
}

function getAllUserQuestionAnswers(): UserQuestionAnswers[] {
  const data = getAllAnswersData();

  const res = data.map((u) => {
    const answers: AnswerWithInlineQuestion[] = u.answers.filter((a) => a.hasOwnProperty("question")).map((a) => {
      // console.log(u);
      const b: AnswerWithInlineQuestion = {
        updatedAt: a.updatedAt,
        questionId: a.question.id,
        category: a.question.category,
        topic: a.question.topic,
      }

      if (typeof a.knowledge !== "undefined") b.knowledge = a.knowledge;
      if (typeof a.motivation !== "undefined") b.motivation = a.knowledge;
      if (typeof a.customScaleValue !== "undefined") b.customScaleValue = a.customScaleValue;

      return b;
    });
    return {
      username: u.username,
      email: u.email,
      formDefinitionID: u.formDefinitionID,
      updatedAt: u.updatedAt,
      answers
    }
  });

  return res as UserQuestionAnswers[];
}

/**
 * Fetches latest categories. Currently hard coded to id of latest catalog
 *
 * @returns
 * @customfunction
 */
function getCategories() {
  const output = getCategoriesData()
    .map((c) => [c.index, c.text, c.id, c.description]);

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
function getQuestions(): any[][] {
  const data: Question[] = getQuestionsData();
  const output = data
    .map((q) => [q.index, q.topic, q.text, q.type, q.id, q.categoryID])
    .sort((a, b) => (a[5] > b[5] ? 1 : -1));
  return output;
}

/**
 * Fetches the questions, merges them with categories and returns as a nested array for use in a spreadsheet
 * 
 * @returns Array<any, any>
 * @customfunction 
 */
function getQuestionsWithCategory() {
  const questions: Question[] = getQuestionsData();
  const categories: Category[] = getCategoriesData();

  const res = questions.map((q) => {
    const cat = categories.find((c) => c.id === q.categoryID);

    const r = [q.index, q.topic, q.text, q.type, q.id, cat?.id, cat?.index, cat?.text, cat?.description];

    return r;
  });

  return res;
}