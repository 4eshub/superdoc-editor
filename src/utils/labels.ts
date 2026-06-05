import type { SuperDocLabels } from '@/types/messages'

const en: SuperDocLabels = {
  openDocxFile: 'Open DOCX file',
  differentFirstPage: 'Different first page (unique header and footer on page 1)',
  differentFirstPageActive: 'Different first page is on. Page 1 has its own header and footer.',
  editingFirstPageHeader: 'Editing first page header',
  editingFirstPageFooter: 'Editing first page footer',
  editingOtherPagesHeader: 'Editing header for other pages',
  editingOtherPagesFooter: 'Editing footer for other pages',
  pageBreak: 'Insert page break',
}

const fr: SuperDocLabels = {
  openDocxFile: 'Ouvrir un fichier DOCX',
  differentFirstPage: 'Première page différente (en-tête et pied de page uniques sur la page 1)',
  differentFirstPageActive:
    'Première page différente activée. La page 1 a son propre en-tête et pied de page.',
  editingFirstPageHeader: "Modification de l'en-tête de la première page",
  editingFirstPageFooter: 'Modification du pied de page de la première page',
  editingOtherPagesHeader: "Modification de l'en-tête des autres pages",
  editingOtherPagesFooter: 'Modification du pied de page des autres pages',
  pageBreak: 'Insérer un saut de page',
}

const ar: SuperDocLabels = {
  openDocxFile: 'فتح ملف DOCX',
  differentFirstPage: 'صفحة أولى مختلفة (ترويسة وتذييل فريدان للصفحة 1)',
  differentFirstPageActive: 'الصفحة الأولى المختلفة مفعّلة. للصفحة 1 ترويسة وتذييل خاصان.',
  editingFirstPageHeader: 'تحرير ترويسة الصفحة الأولى',
  editingFirstPageFooter: 'تحرير تذييل الصفحة الأولى',
  editingOtherPagesHeader: 'تحرير ترويسة الصفحات الأخرى',
  editingOtherPagesFooter: 'تحرير تذييل الصفحات الأخرى',
  pageBreak: 'إدراج فاصل صفحة',
}

const labelsByLocale: Record<string, SuperDocLabels> = {
  en,
  fr,
  ar,
}

export function resolveLabels(locale?: string, overrides?: Partial<SuperDocLabels>): SuperDocLabels {
  const language = locale?.split('-')[0] ?? 'en'
  return {
    ...en,
    ...(labelsByLocale[language] ?? {}),
    ...(overrides ?? {}),
  }
}
