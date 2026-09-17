import fs from 'node:fs/promises';
import decompress from 'woff2-encoder/decompress';
await fs.mkdir('.local/word-default-fonts',{recursive:true});
const data=await fs.readFile('node_modules/@fontsource-variable/inter/files/inter-latin-wght-normal.woff2');
await fs.writeFile('.local/word-default-fonts/Inter-variable.ttf',await decompress(data));
