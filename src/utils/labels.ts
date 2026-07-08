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
  pageNumber: 'Insert page number',
  pageNumberAtCursor: 'Insert page number',
  pageNumberLeft: 'Page number – left',
  pageNumberCenter: 'Page number – center',
  pageNumberRight: 'Page number – right',
  pageNumberNeedsHeaderFooter:
    'Click into the document header or footer first, then insert the page number.',
  pageNumberWrongFirstPageVariant:
    'Different first page is on. Insert page numbers in the header or footer for other pages, not the first page.',
  pageNumberInsertFailed: 'Could not insert the page number. Try clicking into a header or footer first.',
  copyCommentLink: 'Copy comment link',
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
  pageNumber: 'Insérer le numéro de page',
  pageNumberAtCursor: 'Insérer le numéro de page',
  pageNumberLeft: 'Numéro de page – gauche',
  pageNumberCenter: 'Numéro de page – centre',
  pageNumberRight: 'Numéro de page – droite',
  pageNumberNeedsHeaderFooter:
    "Cliquez d'abord dans l'en-tête ou le pied de page du document, puis insérez le numéro de page.",
  pageNumberWrongFirstPageVariant:
    "La première page différente est activée. Insérez les numéros de page dans l'en-tête ou le pied de page des autres pages, pas de la première page.",
  pageNumberInsertFailed:
    "Impossible d'insérer le numéro de page. Essayez de cliquer dans un en-tête ou un pied de page.",
  copyCommentLink: 'Copier le lien du commentaire',
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
  pageNumber: 'إدراج رقم الصفحة',
  pageNumberAtCursor: 'إدراج رقم الصفحة',
  pageNumberLeft: 'رقم الصفحة – يسار',
  pageNumberCenter: 'رقم الصفحة – وسط',
  pageNumberRight: 'رقم الصفحة – يمين',
  pageNumberNeedsHeaderFooter: 'انقر أولاً في ترويسة أو تذييل المستند، ثم أدرج رقم الصفحة.',
  pageNumberWrongFirstPageVariant:
    'الصفحة الأولى المختلفة مفعّلة. أدرج أرقام الصفحات في ترويسة أو تذييل الصفحات الأخرى، وليس الصفحة الأولى.',
  pageNumberInsertFailed: 'تعذّر إدراج رقم الصفحة. جرّب النقر في ترويسة أو تذييل أولاً.',
  copyCommentLink: 'نسخ رابط التعليق',
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
