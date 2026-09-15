import type { PowerPointCanvasJson } from '../src/lib/import/PowerpointImportTypes'

export const VALID_METADATA = {
  subject: {
    summary: ' A security testing architecture with data integrations. ',
    domains: [
      { id: 'cybersecurity' as const, relevance: 'primary' as const, topics: ['Security Testing', 'security-testing'] },
      { id: 'software-architecture' as const, relevance: 'secondary' as const, topics: ['Extensibility'] },
    ],
    other_topics: [],
    technologies: ['Salesforce', 'Snowflake'],
    entities: ['PolicyCenter'],
    claims: ['Testing coverage has a control gap'],
    synonyms: ['Application security assurance'],
  },
  communication: {
    intents: ['finding' as const, 'evidence' as const, 'recommendation' as const],
    information_types: ['Observations', 'Gaps', 'Remediation Actions'],
    audience: ['Technology diligence'],
  },
  template_fit: {
    archetype: 'Finding Evidence Recommendation',
    content_slots: [
      { element_id: 'text-1', role: 'Headline', capacity: 'short-text' as const },
    ],
  },
  visual: {
    layout_type: 'diagram-led' as const,
    visual_elements: ['System Boxes', 'Directional Connectors'],
    content_density: 'medium' as const,
    structural_features: ['Three tiers', 'Legend'],
  },
  retrieval_keywords: ['Application stack', 'System integration'],
}

export function createTemplate(title = 'Product Architecture'): PowerPointCanvasJson {
  return {
    presentation: {
      preserveElementOrder: true,
      showBranding: false,
      title,
      slides: [{
        backgroundColor: 'FFFFFF',
        elements: [
          {
            fill: 'transparent', h: 40, id: 'text-1', stroke: 'transparent', strokeWidth: 0,
            text: 'Customer data flow', type: 'text', w: 400, x: 20, y: 20,
          },
          {
            endArrow: 'triangle', id: 'line-1', lineType: 'straight', stroke: '000000',
            strokeWidth: 1, type: 'line', x1: 100, x2: 300, y1: 200, y2: 200,
          },
          {
            altText: 'Sensitive diagram', fit: 'contain', h: 100, id: 'image-1',
            src: 'data:image/png;base64,DO_NOT_INCLUDE', type: 'image', w: 100, x: 500, y: 200,
          },
        ],
        height: 720,
        id: 'slide-1',
        name: 'Architecture',
        width: 1280,
      }],
    },
  }
}
