import fs from 'fs';
import { sync as which } from 'which';
import YAML from 'yaml';
import { Template } from '../../models';
import { ISession } from '../../session/types';
import { TexExportOptions } from './types';

export function throwIfTemplateButNoJtex(opts: TexExportOptions) {
  if (opts.template && !which('jtex', { nothrow: true })) {
    throw new Error(
      'A template option was specified but the `jtex` command was not found on the path.\nTry `pip install jtex`!',
    );
  }
}

export async function ifTemplateFetchTaggedBlocks(
  session: ISession,
  opts: TexExportOptions,
): Promise<{ tagged: string[] }> {
  let tagged: string[] = [];
  if (opts.template) {
    session.log.debug(`Template: Fetching spec for '${opts.template}'`);
    let requestedTemplate = opts.template.replace(/^tex\//g, '');
    if (requestedTemplate.indexOf('/') === -1) {
      requestedTemplate = `public/${requestedTemplate}`;
      session.log.debug(`Template: Changing from '${opts.template}' to '${requestedTemplate}'`);
    }
    const template = await new Template(session, `tex/${requestedTemplate}`).get();
    tagged = template.data.config.tagged.map((t) => t.id);
    session.log.debug(
      `Template: '${opts.template}' supports following tagged content: ${tagged.join(', ')}`,
    );
  }
  return { tagged };
}

export function ifTemplateLoadOptions(opts: TexExportOptions): Record<string, any> {
  if (opts.options) {
    if (!fs.existsSync(opts.options)) {
      throw new Error(`The template options file specified was not found: ${opts.options}`);
    }
    // TODO validate against the options schema here
    return YAML.parse(fs.readFileSync(opts.options as string, 'utf8')) as Record<string, any>;
  }
  return {};
}
