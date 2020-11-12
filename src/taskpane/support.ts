export function getGlobal() {
  console.log("init globals for command buttons");
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

g.SyncButton = SyncButton;

export function SyncButton(event: Office.AddinCommands.Event) {
  console.log("The SYNC BUTTON WAS PRESSED!");
  event.completed();
}

export function registerAutoSyncEvent() {
  Excel.run(async context => {
    context.workbook.worksheets.onChanged.add(onChange);
    await context.sync().then(function() {
      console.log("Event handler on changed registered!");
    });
  });
}

async function onChange() {
  await Excel.run(async context => {
    enableSyncButton(true);
    await context.sync(function() {
      console.log("Detected change in the document!");
    });
  });
}

export function enableSyncButton(enableSync: boolean = false) {
  // @ts-ignore
  OfficeRuntime.ui
    .getRibbon()
    // @ts-ignore
    .then(ribbon => {
      ribbon.requestUpdate({
        tabs: [
          {
            id: "TabHome",
            controls: [
              {
                id: "TaskpaneButton",
                enabled: true
              },
              {
                id: "SyncButton",
                enabled: enableSync
              }
            ]
          }
        ]
      });
    });
}
