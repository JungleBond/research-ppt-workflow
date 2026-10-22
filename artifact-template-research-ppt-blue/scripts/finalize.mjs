import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
const [buildArg,finalArg,skillArg,pythonArg]=process.argv.slice(2);
if(!pythonArg)throw Error('Usage: node finalize.mjs build-dir final.pptx current-presentations-skill-dir python-executable');
const w=path.resolve(buildArg),final=path.resolve(finalArg),sk=path.resolve(skillArg);
await fs.mkdir(path.dirname(final),{recursive:true});
const requirements=JSON.parse(await fs.readFile(path.join(w,'requirements.json'),'utf8'));
delete requirements.noVideo;
const {finalizePresentation}=await import(pathToFileURL(path.join(sk,'container_tools/artifact_tool_utils.mjs')).href);
const result=await finalizePresentation({
 ...requirements,workspaceDir:path.dirname(w),candidatePath:path.join(w,'candidate.pptx'),finalPath:final,
 pythonExecutable:pythonArg,
 integrityValidatorPath:path.join(sk,'container_tools/inspect_presentation_package_integrity.py'),
 layoutValidatorPath:path.join(sk,'container_tools/inspect_presentation_layout_geometry.py'),
 layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...requirements.requiredNativeTableOwnerSlides.flatMap(n=>['--require-native-table-slide',String(n)])],
 fontPolicy:{basis:'user_request',families:['Microsoft YaHei','微软雅黑']},
 materializeLiteralChartWorkbooks:true,nativeChartTargetApplication:'powerpoint',
 verifyArtifactToolImport:true,receiptPath:path.join(w,'finalization.json')
});console.log(JSON.stringify({finalPath:result.finalPath,passed:true}));
