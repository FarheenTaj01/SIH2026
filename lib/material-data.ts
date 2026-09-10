import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { read, utils } from 'xlsx'

export type Material = {
  Record_ID: number
  CPSE: string
  Material_Code: string
  Material_Description: string
  Category: string
  Material_Type: string
  Material: string
  Grade: string
  Primary_Spec: string
  Secondary_Spec: string
  UOM: string
  Unit_Price_INR: number
  Annual_Quantity: number
  Procurement_Year: number
}

export type GroundTruthPair = { Record_A: number; Record_B: number; Ground_Truth_Label: string; Reason: string }

const normalize = (value: unknown) => String(value ?? '').toLowerCase().replace(/[µ²]/g, (v) => v === 'µ' ? 'u' : '2').replace(/[^a-z0-9]+/g, ' ').trim()
const aliases: Record<string, string> = { ss: 'stainless steel', pcs: 'nos', pc: 'nos', hexagonal: 'hex', dia: 'diameter', swg: 'spiral wound gasket', '3 phase': 'three phase', '3-phase': 'three phase' }
const normalized = (value: unknown) => Object.entries(aliases).reduce((text, [from, to]) => text.replace(new RegExp(`\\b${from}\\b`, 'g'), to), normalize(value))

export function scorePair(a: Material, b: Material) {
  const fields: Array<keyof Material> = ['Category', 'Material_Type', 'Material', 'Grade', 'Primary_Spec', 'Secondary_Spec', 'UOM']
  const matching = fields.filter((field) => normalized(a[field]) && normalized(a[field]) === normalized(b[field]))
  const conflicts = fields.filter((field) => normalized(a[field]) && normalized(b[field]) && normalized(a[field]) !== normalized(b[field]))
  const tokensA = new Set(normalized(a.Material_Description).split(' ').filter(Boolean))
  const tokensB = new Set(normalized(b.Material_Description).split(' ').filter(Boolean))
  const union = new Set([...tokensA, ...tokensB])
  const overlap = union.size ? [...tokensA].filter((token) => tokensB.has(token)).length / union.size : 0
  const structure = matching.length / fields.length
  const score = Math.round((overlap * 0.55 + structure * 0.45) * 100)
  const type = score >= 82 && conflicts.length <= 1 ? 'Likely Equivalent' : score >= 58 ? 'Needs Review' : 'Different'
  return { score, type, matching, conflicts, explanation: `${matching.length} of ${fields.length} structured attributes align; description token overlap is ${Math.round(overlap * 100)}%.` }
}

export async function getMaterialData() {
  const file = await readFile(path.join(process.cwd(), 'data/SIH26099_50row_dataset_with_ground_truth_labels-59b7df.xlsx'))
  const workbook = read(file, { type: 'buffer', cellDates: true })
  const materials = utils.sheet_to_json<Material>(workbook.Sheets.Material_Data, { defval: '' })
  const groundTruth = utils.sheet_to_json<GroundTruthPair>(workbook.Sheets.Ground_Truth_Pairs, { defval: '' })
  const pairs = materials.flatMap((a, index) => materials.slice(index + 1).map((b) => ({ a, b, ...scorePair(a, b) }))).filter((pair) => pair.score >= 58).sort((a, b) => b.score - a.score)
  const predicted = new Set(pairs.filter((pair) => pair.type === 'Likely Equivalent').map((pair) => `${pair.a.Record_ID}-${pair.b.Record_ID}`))
  const truth = new Set(groundTruth.filter((pair) => pair.Ground_Truth_Label === 'EXACT_DUPLICATE' || pair.Ground_Truth_Label === 'EQUIVALENT').map((pair) => `${pair.Record_A}-${pair.Record_B}`))
  const tp = [...truth].filter((key) => predicted.has(key)).length
  const fp = [...predicted].filter((key) => !truth.has(key)).length
  const fn = [...truth].filter((key) => !predicted.has(key)).length
  const tn = Math.max(0, (materials.length * (materials.length - 1)) / 2 - tp - fp - fn)
  const precision = tp + fp ? tp / (tp + fp) : 0
  const recall = tp + fn ? tp / (tp + fn) : 0
  return { materials, groundTruth, pairs, stats: { total: materials.length, cpse: new Set(materials.map((m) => m.CPSE)).size, categories: new Set(materials.map((m) => m.Category)).size, potential: pairs.length, equivalent: predicted.size, underReview: pairs.filter((p) => p.type === 'Needs Review').length, harmonization: Math.round((predicted.size / materials.length) * 100), tp, fp, fn, tn, precision, recall, f1: precision + recall ? (2 * precision * recall) / (precision + recall) : 0 } }
}

export type MaterialData = Awaited<ReturnType<typeof getMaterialData>>
