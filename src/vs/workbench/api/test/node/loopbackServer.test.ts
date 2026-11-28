/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { URI } from '../../../../base/common/uri.js';
import { NullLogger } from '../../../../platform/log/common/log.js';
import { LoopbackAuthServer } from '../../node/loopbackServer.js';

/**
 * Decodes the five entities produced by `htmlAttributeEncodeValue`, the same way a browser
 * does when reading a quoted attribute value back via `getAttribute()`.
 */
function decodeHtmlAttribute(value: string): string {
	return value.replace(/&(lt|gt|quot|apos|amp);/g, (_, entity) => {
		switch (entity) {
			case 'lt': return '<';
			case 'gt': return '>';
			case 'quot': return '"';
			case 'apos': return '\'';
			case 'amp': return '&';
		}
		return _;
	});
}

/**
 * Extracts the source of the redirect anchor tag. Deliberately keys off page text rather than
 * any attribute added by this change, so a regression fails on its own merits.
 */
function getRedirectAnchor(html: string): string {
	const end = html.indexOf('>click here</a>');
	assert.ok(end !== -1, 'redirect anchor should be present');
	const start = html.lastIndexOf('<a ', end);
	assert.ok(start !== -1, 'redirect anchor should have an opening tag');
	return html.substring(start, end + 1);
}

/**
 * Strips quoted attribute values, leaving only the region a browser tokenizes as attribute
 * names. Markup inside a quoted value is inert, so this separates a real handler from an
 * escaped one that merely looks like markup.
 */
function getAnchorAttributeRegion(html: string): string {
	return getRedirectAnchor(html).replace(/"[^"]*"/g, '');
}

/** Extracts the anchor's `href` attribute value, the way a browser parses a quoted attribute. */
function getRedirectHref(html: string): string {
	const match = /href="([^"]*)"/.exec(getRedirectAnchor(html));
	assert.ok(match, 'redirect anchor should have a quoted href attribute');
	return match[1];
}

/** Extracts the trailing inline `<script>` block that performs the redirect. */
function getScriptBlock(html: string): string {
	const start = html.lastIndexOf('<script>');
	const end = html.lastIndexOf('</script>');
	assert.ok(start !== -1 && end > start, 'inline script block should be present');
	return html.substring(start + '<script>'.length, end);
}

function render(appUri: URI, appName = 'Visual Studio Code'): string {
	return new LoopbackAuthServer(new NullLogger(), appUri, appName).getHtml();
}

async function requestCallback(server: LoopbackAuthServer, params: URLSearchParams): Promise<{ status: number | undefined; location: string | undefined }> {
	const { get } = await import('http');
	return new Promise((resolve, reject) => {
		get(`${server.redirectUri}?${params}`, response => {
			response.resume();
			resolve({ status: response.statusCode, location: response.headers.location });
		}).on('error', reject);
	});
}

/**
 * Builds a callback URI whose path segment carries the given authorization server authority,
 * mirroring how the dynamic auth provider constructs it.
 */
function appUriForAuthority(authority: string): URI {
	return URI.from({
		scheme: 'vscode',
		authority: 'dynamicauthprovider',
		path: `/${authority}/redirect`,
		query: 'nonce=NONCE123'
	});
}

suite('LoopbackAuthServer - getHtml', () => {
	ensureNoDisposablesAreLeakedInTestSuite();

	test('does not interpolate the app URI into the inline script', () => {
		const appUri = appUriForAuthority('x\';alert(1)//@evil.example');

		assert.deepStrictEqual(
			{
				scriptMentionsUri: getScriptBlock(render(appUri)).includes('vscode://'),
				scriptMentionsPayload: getScriptBlock(render(appUri)).includes('alert(1)')
			},
			{ scriptMentionsUri: false, scriptMentionsPayload: false }
		);
	});

	test('encodes characters that would break out of the href attribute', () => {
		const appUri = appUriForAuthority('x" onmouseover=alert(1) z@evil.example');

		assert.deepStrictEqual(
			{
				// A raw quote terminates the attribute early and injects an event handler.
				injectsEventHandler: /\son[a-z]+\s*=/i.test(getAnchorAttributeRegion(render(appUri))),
				hasRawQuoteInHref: /["'<>]/.test(getRedirectHref(render(appUri)))
			},
			{ injectsEventHandler: false, hasRawQuoteInHref: false }
		);
	});

	test('href round-trips to the exact app URI', () => {
		// `$waitForUriHandler` compares scheme, authority and path, so the value must survive
		// byte-exact. Production reads `getAttribute`, not `.href`, which normalizes the URL.
		const uris = [
			appUriForAuthority('login.microsoftonline.com'),
			appUriForAuthority('x\';alert(1)//@evil.example'),
			URI.from({
				scheme: 'vscode',
				authority: 'dynamicauthprovider',
				path: '/a"b<c>d&e/redirect',
				query: 'nonce=ABC&state=XYZ'
			})
		];

		assert.deepStrictEqual(
			uris.map(uri => decodeHtmlAttribute(getRedirectHref(render(uri)))),
			uris.map(uri => uri.toString(true))
		);
	});

	test('escapes the app name', () => {
		const html = render(appUriForAuthority('login.example.com'), 'Code <img src=x onerror=alert(1)>');

		assert.deepStrictEqual(
			{
				hasRawTag: html.includes('<img src=x'),
				hasEscapedTag: html.includes('&lt;img src=x onerror=alert(1)&gt;')
			},
			{ hasRawTag: false, hasEscapedTag: true }
		);
	});
});

suite('LoopbackAuthServer - browser callback', () => {
	ensureNoDisposablesAreLeakedInTestSuite();

	test('leaves a pending sign-in usable after callbacks for an earlier sign-in', async () => {
		const server = new LoopbackAuthServer(new NullLogger(), undefined, 'Signal-C Test');
		await server.start();
		try {
			const result = server.waitForOAuthResponse();
			void result.catch(() => { });
			for (const params of [
				new URLSearchParams({ state: 'earlier-state', code: 'earlier-code' }),
				new URLSearchParams({ state: 'earlier-state', error: 'access_denied' }),
				new URLSearchParams({ error: 'access_denied' }),
				new URLSearchParams([['state', server.state], ['state', 'earlier-state'], ['code', 'ambiguous-code']]),
			]) {
				assert.strictEqual((await requestCallback(server, params)).status, 400);
			}
			const current = new URLSearchParams({ state: server.state, code: 'current-code' });
			const response = await requestCallback(server, current);
			assert.strictEqual(response.status, 302);
			assert.strictEqual(response.location, '/done');
			assert.deepStrictEqual(await result, { state: server.state, code: 'current-code' });
			assert.strictEqual((await requestCallback(server, current)).status, 400);
		} finally {
			await server.stop();
		}
	});

	test('returns a denial only for the pending sign-in', async () => {
		const server = new LoopbackAuthServer(new NullLogger(), undefined, 'Signal-C Test');
		await server.start();
		try {
			const result = assert.rejects(server.waitForOAuthResponse(), { message: 'access_denied' });
			const params = new URLSearchParams({ state: server.state, error: 'access_denied', error_description: 'Access denied & returned' });
			const response = await requestCallback(server, params);
			assert.strictEqual(response.status, 302);
			assert.strictEqual(response.location, '/done?error=Access%20denied%20%26%20returned');
			await result;
		} finally {
			await server.stop();
		}
	});
});
