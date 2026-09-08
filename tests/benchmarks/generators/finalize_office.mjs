import path from 'node:path';
import fs from 'node:fs/promises';
import {pathToFileURL,fileURLToPath} from 'node:url';
const base=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const skill=process.env.PRESENTATIONS_SKILL;
if(!skill || !process.env.RUNTIME_PYTHON)throw new Error('Set PRESENTATIONS_SKILL and RUNTIME_PYTHON to the bundled skill/runtime paths.');
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const candidatePath=path.join(base,'generated/scribble-test-kit-v1/evaluator-only/reference-deck.pptx');
const finalPath=path.join(base,'generated/qa/final/reference-deck-final-'+Date.now()+'.pptx');
await fs.mkdir(path.dirname(finalPath),{recursive:true});
const result=await finalizePresentation({workspaceDir:base,candidatePath,finalPath,
 pythonExecutable:process.env.RUNTIME_PYTHON,
 integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),
 layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),
 layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-heading-fit'],
 explicitTotalSlideCount:6,requiredNativeChartOwnerSlides:[2,3],materializeLiteralChartWorkbooks:true,
 fontPolicy:{basis:'design',families:['Arial']},verifyArtifactToolImport:true,receiptPath:path.join(base,'generated/qa/reference-deck-validation-'+Date.now()+'.json')});
await fs.copyFile(finalPath,candidatePath);console.log(JSON.stringify(result));
