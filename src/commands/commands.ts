/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, console */

import { AutoCompleteEngine } from "../autocomplete/engine";

let autoCompleteEngine: AutoCompleteEngine | null = null;

export function setAutoCompleteEngine(engine: AutoCompleteEngine | null): void {
  autoCompleteEngine = engine;
}

// 注册快捷键函数
Office.onReady(() => {
  // 注册快捷键处理函数
  Office.actions.associate("acceptSuggestion", function (event: Office.AddinCommands.Event) {
    console.log("Keyboard shortcut triggered: acceptSuggestion");

    if (autoCompleteEngine) {
      autoCompleteEngine.applySuggestion()
        .then(() => {
          console.log("Suggestion accepted successfully");
          event.completed();
        })
        .catch((error) => {
          console.error("Failed to accept suggestion:", error);
          event.completed();
        });
    } else {
      console.log("AutoCompleteEngine not initialized");
      event.completed();
    }
  });
});
