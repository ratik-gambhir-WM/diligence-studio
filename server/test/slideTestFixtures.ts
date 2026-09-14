import type { PowerPointCanvasJson } from '../src/lib/import/PowerpointImportTypes'

export const VALID_METADATA = {
  slide_type: 'Architecture Overview',
  slide_purpose: ' Explain the platform architecture ',
  description: 'A three-tier platform with integrations.',
  topics: ['Architecture', 'architecture', 'Integrations'],
  business_domains: ['Insurance'],
  technologies: ['Salesforce', 'Snowflake'],
  entities: ['PolicyCenter'],
  use_cases: ['Answer architecture diligence questions'],
  audience: ['Technology diligence'],
  layout_type: 'Diagram Led',
  visual_elements: ['System boxes', 'Directional connectors'],
  content_density: 'medium' as const,
  information_types: ['Current state', 'Dependencies'],
  structural_features: ['Three tiers', 'Legend'],
  retrieval_keywords: ['Application stack', 'System integration'],
  has_timeline: false,
  has_table: false,
  has_chart: false,
  has_process_flow: true,
  has_kpis: false,
  has_recommendations: false,
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
