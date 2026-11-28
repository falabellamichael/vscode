/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { rejects, strictEqual } from 'assert';
import { CancellationToken } from '../../../../../../base/common/cancellation.js';
import { Emitter, Event } from '../../../../../../base/common/event.js';
import { toDisposable } from '../../../../../../base/common/lifecycle.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../../base/test/common/utils.js';
import { TestConfigurationService } from '../../../../../../platform/configuration/test/common/testConfigurationService.js';
import { NullLogService } from '../../../../../../platform/log/common/log.js';
import { GeneralShellType, type ITerminalLogService } from '../../../../../../platform/terminal/common/terminal.js';
import { PromptInputState } from '../../../../../../platform/terminal/common/capabilities/commandDetection/promptInputModel.js';
import { RichExecuteStrategy } from '../../browser/executeStrategy/richExecuteStrategy.js';
import type { ITerminalInstance } from '../../../../terminal/browser/terminal.js';
import type { ICommandDetectionCapability } from '../../../../../../platform/terminal/common/capabilities/capabilities.js';

function createLogService(): ITerminalLogService {
	return new class extends NullLogService { readonly _logBrand = undefined; };
}

suite('RichExecuteStrategy', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();

	test('returns an incomplete PowerShell input to the model and clears the pending editor line', async () => {
		const changed = store.add(new Emitter<void>());
		const command = 'Write-Output "missing quote';
		const marker = { line: 0, dispose() { }, onDispose: Event.None };
		const input = { state: PromptInputState.Input, prefix: command + '\n', suffix: '', continuationPrompt: '>> ', onDidChangeInput: changed.event };
		const buffer = { baseY: 0, cursorY: 1, getLine: () => ({ translateToString: () => '>>' }) };
		let interrupted = '';
		const instance = {
			xtermReadyPromise: Promise.resolve({ raw: { registerMarker: () => marker, buffer: { active: buffer, alternate: {}, onBufferChange: Event.None } } }),
			shellType: GeneralShellType.PowerShell, onData: Event.None, onDisposed: Event.None, onExit: Event.None,
			runCommand: () => changed.fire(), sendText: async (text: string) => { interrupted = text; },
		} as unknown as ITerminalInstance;
		const detection = { promptInputModel: input, onCommandFinished: Event.None } as unknown as ICommandDetectionCapability;
		const strategy = store.add(new RichExecuteStrategy(instance, detection, false, new TestConfigurationService(), createLogService()));
		const result = await strategy.execute(command, CancellationToken.None);
		strictEqual(result.error, 'incompleteCommand');
		strictEqual(interrupted, '\x03');
	});

	test('passes separate command line metadata when running a wrapped command', async () => {
		const onCommandFinishedEmitter = new Emitter<{ getOutput(): string; exitCode: number }>();
		let actualCommandLine: string | undefined;
		let actualCommandId: string | undefined;
		let actualCommandLineForMetadata: string | undefined;

		const marker = {
			line: 0,
			dispose: () => { },
			onDispose: Event.None,
		};
		const xterm = {
			raw: {
				registerMarker: () => marker,
				buffer: {
					active: {},
					alternate: {},
					onBufferChange: () => toDisposable(() => { }),
				},
				getContentsAsText: () => '',
			}
		};
		const instance = {
			xtermReadyPromise: Promise.resolve(xterm),
			onData: Event.None,
			onDisposed: Event.None,
			onExit: Event.None,
			runCommand: (commandLine: string, _shouldExecute: boolean, commandId?: string, _forceBracketedPasteMode?: boolean, commandLineForMetadata?: string) => {
				actualCommandLine = commandLine;
				actualCommandId = commandId;
				actualCommandLineForMetadata = commandLineForMetadata;
				queueMicrotask(() => onCommandFinishedEmitter.fire({ getOutput: () => 'output', exitCode: 0 }));
			},
		} as unknown as ITerminalInstance;
		const commandDetection = {
			onCommandFinished: onCommandFinishedEmitter.event,
		} as unknown as ICommandDetectionCapability;
		const strategy = store.add(new RichExecuteStrategy(
			instance,
			commandDetection,
			false,
			new TestConfigurationService(),
			createLogService(),
		));

		await strategy.execute('sandbox:echo hello', CancellationToken.None, 'tool-command-id', 'echo hello');

		strictEqual(actualCommandLine, 'sandbox:echo hello');
		strictEqual(actualCommandId, 'tool-command-id');
		strictEqual(actualCommandLineForMetadata, 'echo hello');
	});

	test('completes when terminal process exits without shell integration sequences', async () => {
		const onCommandFinishedEmitter = new Emitter<{ getOutput(): string; exitCode: number }>();
		const onExitEmitter = new Emitter<number | undefined>();

		const marker = {
			line: 0,
			dispose: () => { },
			onDispose: Event.None,
		};
		const xterm = {
			raw: {
				registerMarker: () => marker,
				buffer: {
					active: {},
					alternate: {},
					onBufferChange: () => toDisposable(() => { }),
				},
				getContentsAsText: () => 'some output',
			}
		};
		const instance = {
			xtermReadyPromise: Promise.resolve(xterm),
			onData: Event.None,
			onDisposed: Event.None,
			onExit: onExitEmitter.event,
			runCommand: () => {
				// Simulate process exiting without firing onCommandFinished
				queueMicrotask(() => onExitEmitter.fire(1));
			},
		} as unknown as ITerminalInstance;
		const commandDetection = {
			onCommandFinished: onCommandFinishedEmitter.event,
		} as unknown as ICommandDetectionCapability;
		const strategy = store.add(new RichExecuteStrategy(
			instance,
			commandDetection,
			false,
			new TestConfigurationService(),
			createLogService(),
		));

		const result = await strategy.execute('exit 1', CancellationToken.None);

		strictEqual(result.exitCode, 1);
	});

	test('handles ITerminalLaunchError on process exit', async () => {
		const onCommandFinishedEmitter = new Emitter<{ getOutput(): string; exitCode: number }>();
		const onExitEmitter = new Emitter<number | { message: string; code?: number } | undefined>();

		const marker = {
			line: 0,
			dispose: () => { },
			onDispose: Event.None,
		};
		const xterm = {
			raw: {
				registerMarker: () => marker,
				buffer: {
					active: {},
					alternate: {},
					onBufferChange: () => toDisposable(() => { }),
				},
				getContentsAsText: () => '',
			}
		};
		const instance = {
			xtermReadyPromise: Promise.resolve(xterm),
			onData: Event.None,
			onDisposed: Event.None,
			onExit: onExitEmitter.event,
			runCommand: () => {
				queueMicrotask(() => onExitEmitter.fire({ message: 'Failed to launch', code: 127 }));
			},
		} as unknown as ITerminalInstance;
		const commandDetection = {
			onCommandFinished: onCommandFinishedEmitter.event,
		} as unknown as ICommandDetectionCapability;
		const strategy = store.add(new RichExecuteStrategy(
			instance,
			commandDetection,
			false,
			new TestConfigurationService(),
			createLogService(),
		));

		const result = await strategy.execute('bad-command', CancellationToken.None);

		strictEqual(result.exitCode, 127);
	});

	test('returns immediately with captured exit code when pty has already exited before execute()', async () => {
		// Simulates the scenario where the shell process from a previous command
		// has already died, so onExit has already fired and Event.toPromise(onExit)
		// would never resolve. The strategy must short-circuit using the
		// instance's already-captured exitCode.
		const onCommandFinishedEmitter = new Emitter<{ getOutput(): string; exitCode: number }>();
		const onExitEmitter = new Emitter<number | undefined>();
		const instance = {
			xtermReadyPromise: Promise.resolve({}),
			onData: Event.None,
			onDisposed: Event.None,
			onExit: onExitEmitter.event,
			isDisposed: false,
			exitCode: 1,
			runCommand: () => { throw new Error('runCommand should not be called when pty already exited'); },
		} as unknown as ITerminalInstance;
		const commandDetection = {
			onCommandFinished: onCommandFinishedEmitter.event,
		} as unknown as ICommandDetectionCapability;
		const strategy = store.add(new RichExecuteStrategy(
			instance,
			commandDetection,
			false,
			new TestConfigurationService(),
			createLogService(),
		));

		const result = await strategy.execute('Rscript /app/ars.R', CancellationToken.None);

		strictEqual(result.exitCode, 1);
		strictEqual(result.output, undefined);
		strictEqual(result.additionalInformation, 'Command exited with code 1');
	});

	test('throws "The terminal was closed" when instance is already disposed before execute()', async () => {
		const onCommandFinishedEmitter = new Emitter<{ getOutput(): string; exitCode: number }>();
		const instance = {
			xtermReadyPromise: Promise.resolve({}),
			onData: Event.None,
			onDisposed: Event.None,
			onExit: Event.None,
			isDisposed: true,
			exitCode: undefined,
			runCommand: () => { throw new Error('runCommand should not be called when terminal is disposed'); },
		} as unknown as ITerminalInstance;
		const commandDetection = {
			onCommandFinished: onCommandFinishedEmitter.event,
		} as unknown as ICommandDetectionCapability;
		const strategy = store.add(new RichExecuteStrategy(
			instance,
			commandDetection,
			false,
			new TestConfigurationService(),
			createLogService(),
		));

		await rejects(
			() => strategy.execute('echo hello', CancellationToken.None),
			/The terminal was closed/
		);
	});
});
