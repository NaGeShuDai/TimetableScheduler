import { build } from 'esbuild';
const browserOnly={name:'disable-external-native-solver',setup(builder){builder.onResolve({filter:/^(fs|child_process)$/},args=>({path:args.path,namespace:'native-disabled'}));builder.onLoad({filter:/.*/,namespace:'native-disabled'},()=>({contents:'throw new Error("Native external solver is unavailable in browser");'}));}};
await build({entryPoints:['src/app.js'],bundle:true,outfile:'public/app.js',minify:true,platform:'browser',plugins:[browserOnly]});
await build({entryPoints:['src/worker.js'],bundle:true,outfile:'public/worker.js',minify:true,platform:'browser',plugins:[browserOnly]});
