/**
 * Test
 */

function updateCompetencyData() {}

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

function deriveTaxonomy(data: any) {
  const tree = {};

  data[0].answers.forEach((el: Answer) => {
    const q: Question = el.question;
    tree[q.category] = {};
    tree[q.category][q.topic] = q.id;
  });

  return tree;
}

/**
 * Fetches data for competency mapping
 *
 * @returns list of competency data
 * @customfunction
 */
function getCompetencyData() {
  const res = UrlFetchApp.fetch(config.url, {
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
  console.log(data[0].answers);
  const taxonomy = deriveTaxonomy(data);
  console.log(taxonomy);
  const categories = Object.keys(taxonomy).sort();
  categories.unshift('epost', 'timestamp');
  const output = data.map((row: any) => {
    return [row.email, row.updatedAt];
  });

  output.unshift([...categories]);
  console.log(output);
  return output;
}

/**
 * Fetches data and derives only taxonomy (competency catalog)
 *
 * @returns matrix of competencies
 * @customfunction
 */
function getTaxonomy() {
  // foo
}
