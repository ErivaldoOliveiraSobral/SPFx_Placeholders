import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import styles from "./AppPlaceholdersApplicationCustomizer.module.scss";
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'AppPlaceholdersApplicationCustomizerStrings';

const LOG_SOURCE: string = 'AppPlaceholdersApplicationCustomizer';
const topBar: string = 'Esta é a Barra Superior';
const bottomBar: string = 'Esta é a Barra Inferior';
let SUITEBAR: HTMLElement;

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppPlaceholdersApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppPlaceholdersApplicationCustomizer
  extends BaseApplicationCustomizer<IAppPlaceholdersApplicationCustomizerProperties> {

  // These have been added
  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    SUITEBAR = document.getElementById("SuiteNavPlaceHolder") 
                ? document.getElementById("SuiteNavPlaceHolder") 
                : document.querySelector(".od-SuiteNav");

    SUITEBAR.setAttribute("style", "display: none !important")

    document.getElementById("exibirBarra").addEventListener("click", (e: Event) => { 
      this._hideSuiteNavPlaceHolder();
    });

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    console.log(`${LOG_SOURCE}._renderPlaceHolders()`);
    console.log(
      "Available placeholders: ",
      this.context.placeholderProvider.placeholderNames
        .map(name => PlaceholderName[name])
        .join(", ")
    );

    this._getTopPlaceholder();

    this._getBottomPlaceholder();
    
  }

  private _getBottomPlaceholder() {
    // Handling the bottom placeholder
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );

      // The extension should not assume that the expected placeholder is available.
      if (!this._bottomPlaceholder) {
        console.error("The expected placeholder (Bottom) was not found.");
        return;
      }

      if (this.properties) {
        let bottomString: string = bottomBar;
        if (!bottomString) {
          bottomString = "(Bottom property was not defined.)";
        }

        if (this._bottomPlaceholder.domElement) {
          this._bottomPlaceholder.domElement.innerHTML = `
          <div class="${styles.app}">
            <div class="${styles.bottom}">
              ${bottomString}
            </div>
          </div>`;
        }
      }
    }
  }

  private _getTopPlaceholder() {
    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }

      if (this.properties) {
        let topString: string = topBar;
        if (!topString) {
          topString = "(Top property was not defined.)";
        }

        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
          <div class="${styles.app}">
            <div class="${styles.top}">
              <input type="button" id="exibirBarra" value="Exibir Barra">
            </div>
          </div>`;
        }
      }
    }
  }

  private _hideSuiteNavPlaceHolder() {
    var attr: string = SUITEBAR.getAttribute("style");
    if (attr && attr.indexOf("display: none") >= 0) {
      SUITEBAR.removeAttribute("style");
    } else {
      SUITEBAR.setAttribute("style", "display: none !important");
    }
  }

  private _onDispose(): void {
    console.log(`[${LOG_SOURCE}._onDispose] Disposed custom top and bottom placeholders.`);
  }
}
