/**
 * Type definitions
 */
interface Question {
  index: number;
  type: string;
  categoryID: string;
  text: string; // 'Relasjonsdatabaser som Postgres, Oracle o.l.',
  topic: string; // topic: 'Relasjonsdatabaser',
  id: string; // id: 'b1656d08-1f76-443e-a23f-b6179235da75' } },
}

interface AnswerQuestion {
  id: string; // id: 'b1656d08-1f76-443e-a23f-b6179235da75' } },
  topic: string; // topic: 'Relasjonsdatabaser',
  category: string; // category: 'Design',
}

interface AnswerWithNestedQuestion {
  knowledge?: number; // knowledge: 3,
  motivation?: number; //     motivation: 2,
  customScaleValue?: number;
  updatedAt?: string; //     updatedAt: '2021-02-22T12:15:51.688Z',
  question: AnswerQuestion; //     question:
}

interface AnswerWithInlineQuestion {
  knowledge?: number; // knowledge: 3,
  motivation?: number; //     motivation: 2,
  customScaleValue?: number;
  updatedAt?: string; //     updatedAt: '2021-02-22T12:15:51.688Z',
  category: string; // category: 'Design',
  topic: string; // topic: 'Relasjonsdatabaser',
  questionId: string; // id: 'b1656d08-1f76-443e-a23f-b6179235da75' } },
}

interface TaxonomyTree {
  [key: string]: string;
}

interface Category {
  index: number;
  text: string;
  description: string;
  id: string;
}

interface UserAnswers {
  username: string;
  email: string;
  formDefinitionID: string;
  updatedAt: string;
  answers: AnswerWithNestedQuestion[];
}

interface UserQuestionAnswers {
  username: string;
  email: string;
  formDefinitionID: string;
  updatedAt: string;
  answers: AnswerWithInlineQuestion[];
}

interface UserAttribute {
  Name: string;
  Value: any;
}
interface User {
  username: string;
  attributes: UserAttribute[];
}
