import { spawn } from 'child_process';

export type LogLine = (line: string) => void;

export interface RunOptions {
    cwd: string;
    env?: Record<string, string | undefined>;
    log: LogLine;
    /** Run through the system shell (needed for mvn.cmd / npm.cmd on Windows). */
    shell?: boolean;
    /** Text replaced by *** in the log (passwords that could appear in git messages). */
    secrets?: string[];
    timeoutMs?: number;
    /** Command as shown in the log (default: the command and its arguments). */
    display?: string;
    /** false: keep the output out of the log (long listings read by GEIR itself). */
    echo?: boolean;
}

export class CommandError extends Error {
    constructor(message: string, readonly output: string) {
        super(message);
    }
}

/**
 * Runs a command, streams every line to `log` and returns the whole output. Rejects with
 * the last lines when it exits with an error.
 */
export function run(command: string, args: string[], options: RunOptions): Promise<string> {
    return new Promise((resolve, reject) => {
        const hide = (text: string) =>
            (options.secrets ?? []).filter((s) => s && s.length > 2).reduce((t, s) => t.split(s).join('***'), text);
        const shown = options.display ?? (options.shell ? command : [command, ...args].join(' '));
        options.log(`$ ${hide(shown)}`);
        const child = spawn(command, options.shell ? [] : args, {
            cwd: options.cwd,
            env: { ...process.env, ...options.env } as NodeJS.ProcessEnv,
            shell: options.shell ?? false,
            windowsHide: true
        });
        const output: string[] = [];
        let pending = '';
        const onData = (chunk: Buffer) => {
            pending += chunk.toString('utf8');
            const lines = pending.split(/\r?\n/);
            pending = lines.pop() ?? '';
            for (const line of lines) {
                const clean = hide(line.replace(/\x1b\[[0-9;]*m/g, ''));
                output.push(clean);
                if (clean.trim() && options.echo !== false) {
                    options.log(clean);
                }
            }
        };
        child.stdout.on('data', onData);
        child.stderr.on('data', onData);
        const timer = options.timeoutMs
            ? setTimeout(() => {
                  child.kill();
                  reject(new CommandError(`"${hide(shown)}" tardó más de ${Math.round(options.timeoutMs! / 60000)} min y se detuvo.`, output.join('\n')));
              }, options.timeoutMs)
            : null;
        child.on('error', (error: any) => {
            if (timer) clearTimeout(timer);
            const missing = error?.code === 'ENOENT';
            reject(
                new CommandError(
                    missing ? `No se encontró "${command}" en el servidor. Instálalo o indica su ruta en el .env.` : `${command}: ${error.message}`,
                    output.join('\n')
                )
            );
        });
        child.on('close', (code) => {
            if (timer) clearTimeout(timer);
            if (pending.trim()) {
                const clean = hide(pending);
                output.push(clean);
                if (options.echo !== false) {
                    options.log(clean);
                }
            }
            const text = output.join('\n');
            if (code === 0) {
                resolve(text);
            } else {
                const tail = output.filter((l) => l.trim()).slice(-12).join('\n');
                reject(new CommandError(`"${hide(shown)}" terminó con código ${code}.\n${tail}`, text));
            }
        });
    });
}
