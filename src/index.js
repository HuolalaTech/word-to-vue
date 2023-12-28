import * as pkg from 'globby';
// import { exec } from 'node:child_process';
import args from 'args';
import { transform } from './utils.js';
args.option('entry', 'the entry directory', 'word');

const { globby } = pkg;

const argParams = args.parse(process.argv);
// notes:
// 1. only support .docx, and do not support .doc;

// procedures
// 1. translate names, and validate the duplicate template-names
// 2. generate *.vue file
// 3. prettier

async function init() {
  console.log('只支持单个word文件转换');
  const paths = await globby([argParams.entry]);
  // paths.reduce(async (num, filePath) => {
  //   await transform(filePath);
  //   num++;
  //   console.log(num);
  //   return num;
  // }, 0);
  const t = paths[0];
  console.log(`文件所在位置：${t}`);
  transform(t);
}

init();
