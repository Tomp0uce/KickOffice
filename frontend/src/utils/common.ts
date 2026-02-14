import { languageMap } from './constant'

export function getOptionList(map: Record<string, string>, from: 'key' | 'value' = 'key') {
  return from === 'key'
    ? Object.keys(map).map(key => ({
        label: key,
        value: map[key],
      }))
    : Object.values(map).map(key => ({
        label: key,
        value: key,
      }))
}

export const optionLists = {
  localLanguageList: [
    { label: 'English', value: 'en' },
    { label: 'Fran\u00e7ais', value: 'fr' },
  ],
  replyLanguageList: getOptionList(languageMap, 'value'),
}
