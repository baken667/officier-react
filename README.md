# Officier React — experimental controls

Первый слой собственного React-интерфейса для движка ONLYOFFICE: команды, подписки
на состояние выделения и React-панель с bold/italic/undo/redo/save.

**Текущее состояние:** пакет подключается к уже инициализированному `Asc.asc_docs_api`.
Он пока не загружает SDK, не создаёт canvas и не открывает DOCX. Компонента
`<OfficierEditor>` здесь ещё нет. Реальный редактор без iframe пока не реализован.

Основа движка — [Officier / ONLYOFFICE](https://github.com/baken667/officier).
Исходный ONLYOFFICE разработан Ascensio System SIA; Officier — независимая модификация.
Дата создания этого слоя: 2026-09-09. Условия: [LICENSE](LICENSE).

## Использование интерфейса

После установки пакета приложение подключает панель к контроллеру:

```jsx
import {OfficierToolbar} from '@baken667/officier-react';

export function Toolbar({controller}) {
  return <OfficierToolbar controller={controller} labels={{
    toolbar: 'Форматирование', bold: 'Жирный', italic: 'Курсив',
    undo: 'Отменить', redo: 'Повторить', save: 'Сохранить'
  }} />;
}
```

Со стороны runtime (ещё требует реализации загрузки и монтажа SDK):

```js
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
нужна дополнительная переделка SDK. Панель не создаёт iframe и не использует DocsAPI;
это не доказательство готовности всего редактора без iframe.

## Разработка и публикация

```sh
npm ci
npm test
npm pack --dry-run
```

Подготовлен пакет `@baken667/officier-react` версии `0.1.0-alpha.0` и ручной workflow
**Publish alpha package** для GitHub Packages. Публикация не запускается обычным push.
До первого запуска workflow версия в реестре отсутствует. Workflow использует
`GITHUB_TOKEN`, отдельный секрет для публикации не требуется.

После публикации для установки потребуется доступ к пакету в GitHub Packages:

```sh
npm login --scope=@baken667 --auth-type=legacy --registry=https://npm.pkg.github.com
npm install @baken667/officier-react@alpha
```

GitHub требует токен с `read:packages` даже для публичных npm-пакетов.
См. [документацию реестра](https://docs.github.com/en/packages/working-with-a-github-packages-registry/working-with-the-npm-registry).

Unit-тесты используют mock Word API и проверяют команды, доступность операций,
подписки и SSR панели. Они не проверяют загрузку DOCX или работу настоящего SDK.
План runtime и критерии готовности: [ARCHITECTURE.md](ARCHITECTURE.md).
