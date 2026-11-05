import { WorkItemInfo, stripHtml } from '../infrastructure/azure-devops/azure-devops.client';
import type { TaskChanges } from '../infrastructure/gitlab/gitlab.client';
import { FileSummary, technicalDesign } from './change-summary';
import { GeirTask, UAT_RESULTS, emptyTask } from './document';

export interface DraftResult {
    task: GeirTask;
    files: FileSummary[];
}

/** First sentences of a text, cut at a sentence boundary before `max` characters. */
export function summarize(text: string, max = 700): string {
    const flat = text.replace(/\s*\n\s*/g, ' ').replace(/\s+/g, ' ').trim();
    if (flat.length <= max) {
        return flat;
    }
    const cut = flat.slice(0, max);
    const end = Math.max(cut.lastIndexOf('. '), cut.lastIndexOf('; '));
    return end > max * 0.4 ? cut.slice(0, end + 1) : `${cut.trimEnd()}…`;
}

function capitalize(text: string): string {
    return text.charAt(0).toUpperCase() + text.slice(1);
}

/**
 * The benefit is the "para (que) …" part of a user story ("Como X quiero Y para Z");
 * otherwise the first sentence of the description.
 */
export function benefitFrom(description: string): string {
    const flat = description.replace(/\s+/g, ' ').trim();
    const story = /\bpara(?:\s+que)?\s+(.{10,}?)(?:\.\s|$)/i.exec(flat);
    if (story && /\b(quiero|necesito|deseo|como)\b/i.test(flat.slice(0, story.index))) {
        return capitalize(summarize(story[1].replace(/[.;]\s*$/, ''), 400));
    }
    const first = /^(.{10,}?[.!?])(\s|$)/.exec(flat)?.[1] ?? flat;
    return summarize(first, 400);
}

/** Pulls Dado/Cuando/Entonces (or Given/When/Then) out of acceptance criteria. */
export function gherkin(text: string): { given: string; when: string; then: string } | null {
    const flat = text.replace(/\s+/g, ' ');
    const grab = (labels: string, next: string) =>
        new RegExp(`\\b(?:${labels})\\s*:?\\s+(.+?)(?=\\b(?:${next})\\b\\s*:?|$)`, 'i').exec(flat)?.[1]?.trim() ?? '';
    const given = capitalize(grab('Dado(?: que)?|Given', 'Cuando|When|Entonces|Then'));
    const when = capitalize(grab('Cuando|When', 'Entonces|Then|Dado|Given'));
    const then = capitalize(grab('Entonces|Then', 'Dado|Given|Cuando|When'));
    return given || when || then ? { given, when, then } : null;
}

export function buildDraft(id: string, workItem: WorkItemInfo, changes: TaskChanges | null): DraftResult {
    const task = emptyTask(id);
    const description = stripHtml(workItem.descriptionHtml);
    const criteria = stripHtml(workItem.acceptanceCriteriaHtml);

    task.title = workItem.title;
    task.benefit = description ? benefitFrom(description) : workItem.title;
    task.functionalDesign = summarize(description || criteria);
    const tech = technicalDesign(changes?.files ?? []);
    task.technicalDesign = tech.text;
    task.securityDesign = 'N/A';
    task.workItemUrl = workItem.editUrl;

    const images = workItem.discussionImages;
    const evidence = images.find((img) => img.evidenceHint) ?? images[0];
    task.evidence = evidence ? { base64: evidence.base64, extension: evidence.extension } : null;

    const steps = gherkin(criteria) ?? gherkin(description);
    task.testCase = workItem.title;
    task.given = steps?.given || 'Iniciar sesión en el sistema y dirigirse al módulo correspondiente';
    task.when = steps?.when || `Al ejecutar la funcionalidad: ${workItem.title}`;
    task.then = steps?.then || task.benefit;
    task.result = UAT_RESULTS[0];
    task.evidenceLinkText = workItem.taskboardText;
    task.evidenceLinkUrl = workItem.taskboardUrl;

    return { task, files: tech.files };
}
