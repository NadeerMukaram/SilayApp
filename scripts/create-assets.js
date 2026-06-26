const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

function createPng(width, height, r, g, b) {
  const rowSize = 1 + width * 3;
  const paddedRowSize = rowSize + ((4 - (rowSize % 4)) % 4);
  const raw = Buffer.alloc(paddedRowSize * height);

  for (let y = 0; y < height; y += 1) {
    const offset = y * paddedRowSize;
    raw[offset] = 0;
    for (let x = 0; x < width; x += 1) {
      const px = offset + 1 + x * 3;
      raw[px] = r;
      raw[px + 1] = g;
      raw[px + 2] = b;
    }
  }

  const compressed = zlib.deflateSync(raw);

  function chunk(type, data) {
    const length = Buffer.alloc(4);
    length.writeUInt32BE(data.length, 0);
    const typeBuf = Buffer.from(type);
    const crc = Buffer.alloc(4);
    const crcInput = Buffer.concat([typeBuf, data]);
    crc.writeUInt32BE(crc32(crcInput), 0);
    return Buffer.concat([length, typeBuf, data, crc]);
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk('IHDR', ihdr),
    chunk('IDAT', compressed),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

function crc32(buf) {
  let c = ~0;
  for (let i = 0; i < buf.length; i += 1) {
    c ^= buf[i];
    for (let k = 0; k < 8; k += 1) {
      c = (c >>> 1) ^ (0xedb88320 & -(c & 1));
    }
  }
  return ~c >>> 0;
}

const assetsDir = path.join(__dirname, '..', 'assets');
fs.mkdirSync(assetsDir, { recursive: true });

const blue = createPng(1024, 1024, 37, 99, 235);
fs.writeFileSync(path.join(assetsDir, 'icon.png'), blue);
fs.writeFileSync(path.join(assetsDir, 'splash-icon.png'), blue);
fs.writeFileSync(path.join(assetsDir, 'adaptive-icon.png'), blue);
fs.writeFileSync(path.join(assetsDir, 'favicon.png'), createPng(48, 48, 37, 99, 235));

console.log('Assets created in assets/');
