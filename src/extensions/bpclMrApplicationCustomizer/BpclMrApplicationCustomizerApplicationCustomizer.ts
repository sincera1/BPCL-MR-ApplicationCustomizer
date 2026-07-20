import { Log } from '@microsoft/sp-core-library';
import styles from './BpclMrApplicationCustomizer.module.scss';
import Logo from './assets/Bpcl-logo.png';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import Services, { IMenuItem } from './Services';

const LOG_SOURCE =
  'BpclMrApplicationCustomizer';



export default class BpclMrApplicationCustomizerApplicationCustomizer extends BaseApplicationCustomizer<{}> {


  private _top: PlaceholderContent | undefined;
  private _sp!: SPFI;
private _service!: Services;
  private _menuItems: IMenuItem[] = [];

  public async onInit(): Promise<void> {


    Log.info(LOG_SOURCE, 'Initialized');
    this._hideAppBar();
    this._sp = spfi().using(SPFx(this.context));

    // Create Service
    this._service = new Services(this._sp);

    // Load Menu from SharePoint
    this._menuItems = await this._service.getMasterMenu();

    await this._renderTop();

    this._renderBottom();

    this.context.application.navigatedEvent.add(
      this,
      () => {

        this._renderBottom();

      }
    );

    return Promise.resolve();

  }

  private _hideAppBar():
    void {

    const style =
      document.createElement(
        'style'
      );

    style.innerHTML = `

      #spSiteHeader {
        display: none !important;
      }

      #sp-appBar {
        display: none !important;
      }

      .ms-HorizontalNav {
        margin-left: 0 !important;
      }

      #spCommandBar {
        display: none !important;
      }

      #vpc_Page\\.SiteFooter\\.internal\\.03025612-a400-4804-a78e-e1493200a43b {
        display: none !important;
      }

      #CommentsWrapper {
        display: none !important;
      }

    `;

    document.head.appendChild(
      style
    );

  }

  private async _renderTop(): Promise<void> {


    if (document.getElementById('bpcl-custom-navbar')) {
      return;
    }


    this._top = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);


    if (
      !this._top ||
      !this._top.domElement
    ) {
      return;
    }

    const selectedLanguage =
      localStorage.getItem("MRLanguage") || "English";

    let menuHtml = "";

    this._menuItems.forEach((item: IMenuItem) => {

      let title = item.LinkTitle;

      if (selectedLanguage === "Hindi") {
        title = item.TitleHindi;
      }
      else if (selectedLanguage === "Marathi") {
        title = item.TitleMarathi;
      }

      // Handle Hyperlink column
      const url =
        item.RedirectURL && typeof item.RedirectURL === "object"
          ? item.RedirectURL.Url
          : (item.RedirectURL );

      menuHtml += `
    <a href="${url}" class="${styles.navLink}">
      ${title}
    </a>
  `;
    });

    this._top.domElement.innerHTML = `

<div id="bpcl-custom-navbar" class="${styles.mainNavbar}">

<div class="${styles.navbarContainer}">

<div class="${styles.logoSection}">

<img src="${Logo}" alt="BPCL Logo" class="${styles.logoImg}" />

<span class="${styles.logoText}">MR Connect</span>

</div> 

<button
      id="bpclMenuBtn"
      class="${styles.menuBtn}"
      aria-label="Toggle Menu">
      ☰
    </button>

<div id="bpclMenu" class="${styles.navMenu}">

<div class="${styles.navLinks}">
  ${menuHtml}
</div>

<div class="${styles.languageWrapper}">

<select id="mrLanguageSelector" class="${styles.languageSelect}">

<option
value="English"
${selectedLanguage === "English"
        ? "selected"
        : ""}
>
English
</option>

<option
value="Hindi"
${selectedLanguage === "Hindi"
        ? "selected"
        : ""}
>
Hindi
</option>

<option
value="Marathi"
${selectedLanguage === "Marathi"
        ? "selected"
        : ""}
>
Marathi
</option>

</select>

</div>

</div>

</div>

</div>

`;

    setTimeout(() => {

      const menuBtn =
        document.getElementById('bpclMenuBtn');

      const menu =
        document.getElementById('bpclMenu');

      if (menuBtn && menu) {

        menuBtn.addEventListener('click', () => {

          const isOpen =
            menu.classList.contains('bpcl-mobile-open');

          if (isOpen) {

            menu.classList.remove('bpcl-mobile-open');
            menuBtn.textContent = '☰';

          } else {

            menu.classList.add('bpcl-mobile-open');
            menuBtn.textContent = '✕';

          }

        });

      }

    }, 300);

    const dropdown =
      document.getElementById(
        "mrLanguageSelector"
      ) as HTMLSelectElement;

    if (dropdown) {

      dropdown.addEventListener(
        "change",
        async (e: Event) => {

           const selectedValue = (e.target as HTMLSelectElement).value;

          localStorage.setItem(
            "MRLanguage",
            selectedValue
          );

          if (this._top) {

            this._top.dispose();

            this._top = undefined;

          }

          await this._renderTop();

          window.dispatchEvent(
            new CustomEvent(
              "languageChanged"
            )
          );

        }
      );

    }

  }

  private _renderBottom():
    void {

    if (
      document.getElementById(
        'bpcl-footer'
      )
    ) {
      return;
    }

    const pageCanvas =
      document.querySelector(
        '#spPageCanvasContent'
      );

    if (!pageCanvas) {
      return;
    }

    const year =
      new Date()
        .getFullYear();

    const footer =
      document.createElement(
        'div'
      );

    footer.id =
      'bpcl-footer';

    footer.innerHTML = `

<div class="${styles.footerSection}">

<div class="${styles.footerContent}">

<div class="${styles.footerLeft}">
© ${year}
Bharat Petroleum Corporation Limited
</div>

<div class="${styles.footerRight}">
<a href="#">Feedback</a>
<a href="#">Help</a>
</div>

</div>

</div>

`;

    pageCanvas.appendChild(
      footer
    );

  }

}

