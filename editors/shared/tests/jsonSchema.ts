/**
 * 测试用的 JSON Schema 子集校验器（不引入依赖）：docs/wbp/schema 里的 Schema 只用这个子集。
 * 支持：type const enum pattern min/max length/imum maxItems items properties required additionalProperties
 * patternProperties propertyNames $ref($defs)。返回错误列表，空 = 通过。
 */
import { readFileSync } from 'fs';
import { join } from 'path';

type Json = Record<string, unknown>;
const dir = join(__dirname, '..', '..', '..', 'docs', 'wbp', 'schema');
export const load = (n: string): Json => JSON.parse(readFileSync(join(dir, n), 'utf8')) as Json;

/** 返回错误列表（空 = 通过）。支持：type const enum pattern min/max length/imum items properties required additionalProperties patternProperties propertyNames $ref($defs) */
export function validate(schema: Json, value: unknown, root: Json = schema, path = '$'): string[] {
  const errs: string[] = [];
  if (typeof schema.$ref === 'string') {
    const target = (root.$defs as Record<string, Json>)[schema.$ref.replace('#/$defs/', '')];
    return validate(target, value, root, path);
  }
  if ('const' in schema && value !== schema.const) errs.push(`${path}: 应为 ${JSON.stringify(schema.const)}`);
  if (Array.isArray(schema.enum) && !schema.enum.includes(value)) errs.push(`${path}: 不在枚举里 ${JSON.stringify(value)}`);
  const t = schema.type as string | undefined;
  const typeOk = (): boolean => {
    switch (t) {
      case 'object':
        return !!value && typeof value === 'object' && !Array.isArray(value);
      case 'array':
        return Array.isArray(value);
      case 'string':
        return typeof value === 'string';
      case 'number':
        return typeof value === 'number';
      case 'integer':
        return Number.isInteger(value);
      case 'boolean':
        return typeof value === 'boolean';
      default:
        return true;
    }
  };
  if (t && !typeOk()) return [...errs, `${path}: 类型应为 ${t}`];
  if (typeof value === 'string') {
    if (typeof schema.pattern === 'string' && !new RegExp(schema.pattern, 'u').test(value)) errs.push(`${path}: 不匹配 ${schema.pattern}`);
    if (typeof schema.minLength === 'number' && value.length < schema.minLength) errs.push(`${path}: 太短`);
    if (typeof schema.maxLength === 'number' && value.length > schema.maxLength) errs.push(`${path}: 太长`);
  }
  if (typeof value === 'number') {
    if (typeof schema.minimum === 'number' && value < schema.minimum) errs.push(`${path}: 小于 ${schema.minimum}`);
    if (typeof schema.maximum === 'number' && value > schema.maximum) errs.push(`${path}: 大于 ${schema.maximum}`);
  }
  if (Array.isArray(value)) {
    if (typeof schema.maxItems === 'number' && value.length > schema.maxItems) errs.push(`${path}: 元素太多`);
    if (schema.items) value.forEach((v, i) => errs.push(...validate(schema.items as Json, v, root, `${path}[${i}]`)));
  }
  if (value && typeof value === 'object' && !Array.isArray(value)) {
    const obj = value as Json;
    const props = (schema.properties ?? {}) as Record<string, Json>;
    const pats = Object.keys((schema.patternProperties ?? {}) as Json).map((p) => new RegExp(p));
    for (const r of (schema.required ?? []) as string[]) if (!(r in obj)) errs.push(`${path}: 缺少 ${r}`);
    for (const [k, v] of Object.entries(obj)) {
      if (props[k]) errs.push(...validate(props[k], v, root, `${path}.${k}`));
      else if (pats.some((re) => re.test(k))) continue;
      else if (schema.additionalProperties === false) errs.push(`${path}: 不允许的字段 ${k}`);
      else if (schema.additionalProperties && typeof schema.additionalProperties === 'object') errs.push(...validate(schema.additionalProperties as Json, v, root, `${path}.${k}`));
      const names = schema.propertyNames as Json | undefined;
      if (names && typeof names.maxLength === 'number' && k.length > names.maxLength) errs.push(`${path}: 字段名太长 ${k}`);
    }
  }
  return errs;
}

