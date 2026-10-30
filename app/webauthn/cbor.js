/**
 * Minimal CBOR (RFC 8949) encoder for the WebAuthn attestation object.
 * Supports integers, byte strings, text strings, arrays, string-keyed maps,
 * booleans and null. Replaces cbor-x, whose native cbor-extract dependency
 * breaks electron-builder install-app-deps on Electron 43.
 */

function encodeHead(major, n) {
	if (n < 24) return Buffer.from([(major << 5) | n]);
	if (n < 0x100) return Buffer.from([(major << 5) | 24, n]);
	if (n < 0x10000) return Buffer.from([(major << 5) | 25, n >> 8, n & 0xff]);
	if (n < 0x100000000) {
		const b = Buffer.alloc(5);
		b[0] = (major << 5) | 26;
		b.writeUInt32BE(n, 1);
		return b;
	}
	const b = Buffer.alloc(9);
	b[0] = (major << 5) | 27;
	b.writeUInt32BE(Math.floor(n / 0x100000000), 1);
	b.writeUInt32BE(n % 0x100000000, 5);
	return b;
}

function encodeCbor(value) {
	if (value === null) return Buffer.from([0xf6]);
	if (value === true) return Buffer.from([0xf5]);
	if (value === false) return Buffer.from([0xf4]);
	if (typeof value === 'number') {
		if (!Number.isInteger(value)) throw new TypeError('CBOR: only integers are supported');
		return value >= 0 ? encodeHead(0, value) : encodeHead(1, -1 - value);
	}
	if (typeof value === 'string') {
		const bytes = Buffer.from(value, 'utf8');
		return Buffer.concat([encodeHead(3, bytes.length), bytes]);
	}
	if (value instanceof Uint8Array) {
		return Buffer.concat([encodeHead(2, value.length), Buffer.from(value)]);
	}
	if (Array.isArray(value)) {
		return Buffer.concat([encodeHead(4, value.length), ...value.map(encodeCbor)]);
	}
	if (typeof value === 'object') {
		const keys = Object.keys(value);
		const parts = [encodeHead(5, keys.length)];
		for (const key of keys) parts.push(encodeCbor(key), encodeCbor(value[key]));
		return Buffer.concat(parts);
	}
	throw new TypeError(`CBOR: unsupported type ${typeof value}`);
}

module.exports = { encodeCbor };
