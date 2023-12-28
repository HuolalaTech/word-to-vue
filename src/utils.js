import mammoth from 'mammoth';
import fs from 'fs-extra';
import prettier from 'prettier';
const templateNameMap = new WeakMap();

let str = '';
let num = 0;

const numberEnum = {
  1: '一',
  2: '二',
  3: '三',
  4: '四',
  5: '五',
  6: '六',
  7: '七',
  8: '八',
  9: '九',
  10: '十',
};

const transformChildren = (
  children = [],
  isCenter,
  num,
  isIndent = true
) => {
  let childrenStr = '';
  let titleClassName = 'doc-normal';
  children.forEach((child, index) => {
    const boldClass = child.isBold ? 'bold ' : '';
    // 居中 做标题处理
    if (isCenter) {
      if (child.fontSize == 16) {
        titleClassName = 'page-title center bold';
      } else {
        titleClassName = 'sub-title center bold';
      }
    }
    // 字体适配
    const fontSizeClass =
      child.fontSize && !isCenter ? `font-size-${child.fontSize} ` : '';
    const classnames = boldClass + fontSizeClass;
    // 文本：空格 超链接处理
    const text = (child.children || []).reduce((sttr, txt) => {
      if (txt && txt.type === 'text' && txt.value) {
        sttr += txt.value.replace(/(\s)/gi, '&nbsp;');
      }
      if (
        txt &&
        txt.type === 'hyperlink' &&
        txt.children &&
        txt.children.length > 0
      ) {
        sttr += `<a target="_blank" class="blue" href="${txt.href}">${txt.children[0].value}</a>`;
      }
      return sttr;
    }, '');
    // 对序列进行处理，但是无法处理使用word中的排列功能
    const numStr = handleNumberCh(num);
    const span = `<span class="${classnames}">${
      num > 0 && index === 0 ? `${numStr}、  ` : ''
    }${text}</span>`;
    childrenStr += span;
  });
  const indentTab = isIndent ? '&nbsp;&nbsp;&nbsp;&nbsp;' : '';
  return `<p class="${titleClassName}">${indentTab}${childrenStr}</p>\n`;
}

// 读取word路径名，并进行翻译，自动生成template name和文件名
function getTranslatedNames(path) {
  const chArr = path.replace('.docx', '').split('/');
  return chArr[chArr.length - 1];
  // 可以将文件名翻译为英文
  // return new Promise((resolve) => {
  //   const chArr = path.replace('.docx', '').split('/');
  //   exec(
  //     `trans -t english '${chArr[chArr.length - 1]}'`,
  //     (error, stdout, stderr) => {
  //       const arr = stdout.split('\n');
  //       let target = arr[3];
  //       target = target
  //         .replace(/\[1m|\[22m/gi, '')
  //         .split(' ')
  //         .slice(-5)
  //         .join('-');
  //       if (templateNameMap.has(target)) {
  //         let num = templateNameMap.get(target);
  //         target = `${target}${num}`;
  //         templateNameMap.set(target, num++);
  //       }
  //       console.log('error', error);
  //       resolve(target.replace(/[^a-zA-Z-]/gi, ''));
  //     }
  //   );
  // });
}

// 针对word文档中的排列功能，对1 2 3 4 5 做的转换
function handleNumberCh(number) {
  if (Number(number) > 0) {
    if (Number(number) <= 10) {
      return numberEnum[number];
    } else {
      return (number / 10)
        .toString()
        .split('.')
        .reduce((str, item, index) => {
          if (index === 0) {
            str += `${item > 1 ? numberEnum[item] : ''}十`;
          } else {
            str += `${numberEnum[item]}`;
          }

          return str;
        }, '');
    }
  }
  return number;
}

// 处理段落
function handleParagraph(
  node,
  options = {
    isIndent: true,
  }
) {
  let paragraphStr = '';
  if (node.children && node.children.length > 0) {
    const isCenter = node.alignment === 'center';
    if (node.numbering) {
      num++;
    } else {
      num = 0;
    }
    paragraphStr += transformChildren(
      node.children,
      isCenter,
      num,
      options.isIndent
    );
  } else {
    // 换行处理
    paragraphStr += '<br/>';
  }
  return paragraphStr;
}

// 处理table表单
function handleTable(node) {
  let tableStr = `<table border="1"
  cellspacing="0"
  cellpadding="0">`;
  let tbody = '<tbody>';
  if (node.children && node.children.length > 0) {
    node.children.forEach((ele, index) => {
      // 第一列 表头
      if (index === 0) {
        const arr = [];
        (ele.children || []).forEach((the) => {
          const theadStr = the.children.reduce((str, item) => {
            str += handleParagraph(item, {
              isIndent: false,
            });
            return str;
          }, '');
          arr.push(theadStr);
        });
        const thead = `<thead>
        <th width="20%">${arr[0]}</th>
        <th width="80%">${arr[1]}</th>
      </thead>`;
        tableStr += thead;
      } else {
        const tbodyArr = [];
        (ele.children || []).forEach((the) => {
          const tbodyStr = the.children.reduce((str, item) => {
            str += handleParagraph(item, {
              isIndent: false,
            });
            return str;
          }, '');
          tbodyArr.push(tbodyStr);
        });
        const ttr = `<tr>
          <td width="20%">${tbodyArr[0]}</td>
          <td width="80%">${tbodyArr[1]}</td>
        </tr>`;
        tbody += ttr;
        if (index === node.children.length - 1) {
          tbody += '</tbody>';
        }
      }
    });
  }
  tableStr += tbody;
  tableStr += '</table>';
  str += tableStr;
}

const transformElement = (element) => {
  // 根节点
  const document = element.children;
  if (Array.isArray(document)) {
    document.forEach((ele) => {
      switch (ele.type) {
        case 'paragraph':
          str += handleParagraph(ele);
          break;
        case 'table':
          handleTable(ele);
          break;
        default:
          break;
      }
    });
  }

  return element;
}

const options = {
  transformDocument: transformElement,
};

export const transform = (path) => {
  return new Promise((resolve) => {
    mammoth
      .convertToHtml(
        {
          path,
        },
        options
      )
      .done(async () => {
        const templateName = await getTranslatedNames(path);
        // 生成template模版
        const originalStr = `<template>
          <!-- 文件路径： ${path} -->
          <div class="${templateName} page-demo">\n${str}</div>
          </template>
        <script>
        export default {
        name: '${templateName}',
        data () {
            return {
            };
        }
        };
        </script>
          <style lang="less" scoped>
          @import './assets/css/rules.less';
          .bold {
              font-weight: 700;
          }
          .font-size-16 {
              font-size: 16px;
          }
          .font-size-14 {
            font-size: 14px;
          }
          .font-size-12 {
            font-size: 12px;
          }
          .blue {
            color: blue;
          }
          </style>\n`;
        // 格式化
        const finalStr = await prettier.format(originalStr, {
          parser: 'vue', // 把双引号换成单引号
          singleQuote: true,
          // 在代码尾部添加分号
          semi: true,
          printWidth: 500,
        });

        fs.writeFileSync(`${templateName}.vue`, finalStr, 'utf-8');
        console.log('done');
        resolve();
      });
  });
};
