# Officier React — experimental direct editor

Первый слой собственного React-интерфейса для движка ONLYOFFICE: команды, подписки
на состояние выделения, React-панель с bold/italic/undo/redo/save и typed runtime
loader для прямого SDK-монтажа.

**Текущее состояние:** `<OfficierEditor>` уже владеет React lifecycle, загружает
`runtime/manifest.json`, подключает runtime assets DocumentServer и вызывает
`window.OfficierDirectRuntime.mountWord(...)`. Officier DocumentServer отдаёт
thin direct bridge: он загружает socket.io + SDKJS, создаёт `Asc.asc_docs_api`
в DOM-контейнере React-приложения и подаёт WOPI bootstrap в `Asc.asc_CDocInfo`.
Это экспериментальный путь; browser smoke пока проверяет загрузку direct runtime
assets без iframe, а полный DOCX open/edit/save без iframe остаётся следующим
подтверждаемым рубежом.

Основа движка — [Officier / ONLYOFFICE](https://github.com/baken667/officier).
Исходный ONLYOFFICE разработан Ascensio System SIA; Officier — независимая модификация.
Дата создания этого слоя: 2026-09-09. Условия: [LICENSE](LICENSE).

## Использование интерфейса

Минимальный lifecycle редактора:

```tsx
import {OfficierEditor, createOfficierWopiSession} from '@baken667/officier-react';

const session = await createOfficierWopiSession({
  documentServerUrl: '/officier/',
  wopiSrc: 'https://app.example/wopi/files/1',
  accessToken: 'short-lived-token-from-backend',
  accessTokenTtl: Date.now() + 60 * 60 * 1000,
  lang: 'ru'
});

export function Editor() {
  return (
    <OfficierEditor
      documentServerUrl="/officier/"
      session={session}
      onReady={({editor}) => editor.controller.setSession({ready: true, readOnly: false})}
      onError={({code, error}) => console.error(code, error)}
    />
  );
}
```

`documentServerUrl="/officier/"` означает, что пакет запросит
`/officier/runtime/manifest.json`. Manifest перечисляет JS assets прямого
runtime: socket.io, `sdkjs/word/sdk-all-min.js` и `runtime/direct-word-adapter.js`.
Загруженный runtime предоставляет `window.OfficierDirectRuntime` с методом
`mountWord(...)`, который монтирует низкоуровневый Word SDK в переданный DOM-узел.

`createOfficierWopiSession()` вызывает серверный endpoint
`POST /officier/sessions/wopi/word/edit?wopisrc=...` и получает JSON bootstrap,
переиспользуя ту же WOPI-подготовку, что и стандартная ONLYOFFICE host page.

После установки пакета приложение подключает панель к контроллеру:

```tsx
import {OfficierToolbar} from '@baken667/officier-react';

export function Toolbar({controller}) {
  return <OfficierToolbar controller={controller} labels={{
    toolbar: 'Форматирование', bold: 'Жирный', italic: 'Курсив',
    undo: 'Отменить', redo: 'Повторить', save: 'Сохранить'
  }} />;
}
```

Для низкоуровневого runtime adapter:

```ts
import {createWordController} from '@baken667/officier-react/controller';

// api — уже работающий экземпляр Asc.asc_docs_api, полученный от вашего runtime.
const controller = createWordController(api);
// После открытия документа и определения фактических прав:
controller.setSession({ready: true, readOnly: false});
// При переходе в view/потере права редактировать:
controller.setSession({ready: true, readOnly: true});
// При закрытии UI:
controller.dispose();
```

Контроллер создаётся в lifecycle runtime, а не в функции React render. Его dispose
снимает только собственные подписки; освобождением движка, DOM и сокетов занимается
runtime. Состояние readOnly управляет UI и не заменяет авторизацию сервера.
Вызов save означает отправку команды SDK, а не подтверждение сохранения на сервере.
`useWordState(controller)` позволяет строить полностью свои панели.

Нужен один word-engine на страницу: текущий upstream использует глобальные
`window.editor`, `Asc.editor` и общую таблицу callback. Для нескольких редакторов
нужна дополнительная переделка SDK. Пакет не использует DocsAPI iframe path; runtime
проверяет, что direct mount не оставил iframe внутри контейнера редактора.

## Разработка и публикация

```sh
pnpm install --frozen-lockfile
pnpm test
pnpm pack:dry
```

Пакет `@baken667/officier-react@0.1.0-alpha.0` опубликован в GitHub Packages
([workflow публикации](https://github.com/baken667/officier-react/actions/runs/34324910219)).
Для следующих версий увеличьте version и запустите ручной workflow
**Publish alpha package**. Публикация не запускается обычным push. Workflow использует
`pnpm` и `GITHUB_TOKEN`, отдельный секрет для публикации не требуется.

Исходники пакета написаны на TypeScript. Публикуются собранные файлы из `dist`,
а декларации `.d.ts` генерируются компилятором.

Для установки потребуется доступ к пакету в GitHub Packages:

```sh
pnpm config set @baken667:registry https://npm.pkg.github.com
pnpm add @baken667/officier-react@alpha
```

Для Bun-проекта:

```sh
bun add @baken667/officier-react@alpha
```

GitHub требует токен с `read:packages` даже для публичных npm-пакетов.
См. [документацию реестра](https://docs.github.com/en/packages/working-with-a-github-packages-registry/working-with-the-npm-registry).

Unit-тесты используют mock Word API и проверяют команды, доступность операций,
подписки, manifest loader, single-mount guard и SSR панели. Серверный smoke
дополнительно проверяет загрузку direct runtime assets в браузере без iframe.
Полное открытие DOCX через direct bridge ещё нужно закрепить отдельным browser
тестом. План runtime и критерии готовности:
[ARCHITECTURE.md](ARCHITECTURE.md).
