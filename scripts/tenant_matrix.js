#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

function getArgValue(flag) {
  const index = process.argv.indexOf(flag);
  if (index === -1 || index + 1 >= process.argv.length) return null;
  return process.argv[index + 1];
}

function hasFlag(flag) {
  return process.argv.includes(flag);
}

function fail(message) {
  console.error(message);
  process.exit(1);
}

function loadRegistry(registryPath) {
  try {
    return JSON.parse(fs.readFileSync(registryPath, 'utf8'));
  } catch (error) {
    fail(`Failed to read tenant registry at ${registryPath}: ${error.message}`);
  }
}

function validateTenant(tenant, index) {
  if (!tenant || typeof tenant !== 'object') {
    fail(`Tenant entry at index ${index} must be an object.`);
  }

  const requiredFields = ['name', 'stage'];
  requiredFields.forEach(field => {
    if (!tenant[field] || typeof tenant[field] !== 'string') {
      fail(`Tenant "${tenant.name || index}" is missing required string field "${field}".`);
    }
  });

  const hasDirectScriptId = typeof tenant.scriptId === 'string' && tenant.scriptId.trim() !== '';
  const hasScriptIdVar = typeof tenant.scriptIdVar === 'string' && tenant.scriptIdVar.trim() !== '';

  if (!hasDirectScriptId && !hasScriptIdVar) {
    fail(`Tenant "${tenant.name}" must define either "scriptId" or "scriptIdVar".`);
  }

  if (!['dev', 'prod'].includes(tenant.stage)) {
    fail(`Tenant "${tenant.name}" has invalid stage "${tenant.stage}". Use "dev" or "prod".`);
  }

  if (typeof tenant.runRemoteTests !== 'boolean') {
    fail(`Tenant "${tenant.name}" must include boolean field "runRemoteTests".`);
  }
}

function filterTenants(tenants, stage, testsOnly) {
  return tenants.filter(tenant => {
    if (stage !== 'all' && tenant.stage !== stage) return false;
    if (testsOnly && !tenant.runRemoteTests) return false;
    return true;
  });
}

function main() {
  const registryPath = path.resolve(process.cwd(), getArgValue('--registry') || 'deploy/tenants.json');
  const requestedStage = getArgValue('--stage') || 'all';
  const testsOnly = hasFlag('--tests-only');
  const validateOnly = hasFlag('--validate');
  const outputFormat = getArgValue('--format') || 'array';

  if (!['dev', 'prod', 'all'].includes(requestedStage)) {
    fail(`Invalid --stage "${requestedStage}". Use dev, prod, or all.`);
  }

  if (!['array', 'matrix'].includes(outputFormat)) {
    fail(`Invalid --format "${outputFormat}". Use array or matrix.`);
  }

  const registry = loadRegistry(registryPath);
  if (!registry || !Array.isArray(registry.tenants) || !registry.tenants.length) {
    fail(`Tenant registry ${registryPath} must contain a non-empty "tenants" array.`);
  }

  registry.tenants.forEach(validateTenant);

  if (validateOnly) {
    console.log(`Validated ${registry.tenants.length} tenant definition(s) from ${registryPath}.`);
    return;
  }

  const filtered = filterTenants(registry.tenants, requestedStage, testsOnly);
  if (outputFormat === 'matrix') {
    console.log(JSON.stringify({ include: filtered }));
    return;
  }

  console.log(JSON.stringify(filtered));
}

main();
