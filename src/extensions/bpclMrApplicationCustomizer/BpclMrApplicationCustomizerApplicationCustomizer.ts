
import { Log } from '@microsoft/sp-core-library';
import styles from './BpclMrApplicationCustomizer.module.scss';
import Logo from './assets/MRIonnectlogo.png';

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

const LOG_SOURCE = 'BpclMrApplicationCustomizer';

export default class BpclMrApplicationCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<{}> {

  private _top: PlaceholderContent | undefined;
  private _sp!: SPFI;
  private _service!: Services;
  private _menuItems: IMenuItem[] = [];
  private _clockInterval: number | undefined;

  public async onInit(): Promise<void> {

    Log.info(LOG_SOURCE, 'Initialized');

    this._hideAppBar();

    this._sp = spfi().using(SPFx(this.context));

    this._service = new Services(this._sp);

    this._menuItems = await this._service.getMasterMenu();

    await this._renderTop();

    this._startDateTimeClock();

    this._renderBottom();

    this.context.application.navigatedEvent.add(
      this,
      () => {
        this._renderBottom();
      }
    );

    return Promise.resolve();
  }

  private _getCurrentDateTime(): string {

    const now = new Date();

    const day = String(now.getDate()).padStart(2, '0');

    const monthNames = [
      'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
      'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'
    ];

    const month = monthNames[now.getMonth()];
    const year = now.getFullYear();

    let hours = now.getHours();

    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');

    const ampm = hours >= 12 ? 'PM' : 'AM';

    hours = hours % 12;
    hours = hours ? hours : 12;

    const formattedHours = String(hours).padStart(2, '0');

    return `${day} ${month} ${year} ${formattedHours}:${minutes}:${seconds} ${ampm}`;
  }

  private _startDateTimeClock(): void {

    if (this._clockInterval !== undefined) {
      window.clearInterval(this._clockInterval);
    }

    this._clockInterval = window.setInterval(() => {

      const dateTimeElement =
        document.getElementById('mrCurrentDateTime');

      if (dateTimeElement) {
        dateTimeElement.textContent =
          this._getCurrentDateTime();
      }

    }, 1000);
  }

  private _hideAppBar(): void {

    const style = document.createElement('style');

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

    document.head.appendChild(style);
  }

  
private async _renderTop(): Promise<void> {

  if (document.getElementById('bpcl-custom-navbar')) {
    return;
  }

  this._top =
    this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top
    );

  if (!this._top || !this._top.domElement) {
    return;
  }

  const selectedLanguage =
    localStorage.getItem('MRLanguage') || 'English';

  // Get the appropriate menu title for the selected language.
  const getMenuTitle = (item: IMenuItem): string => {

    if (selectedLanguage === 'Hindi') {
      return item.TitleHindi || item.LinkTitle || '';
    }

    if (selectedLanguage === 'Marathi') {
      return item.TitleMarathi || item.LinkTitle || '';
    }

    return item.LinkTitle || '';
  };

  // Escape dynamic values before inserting them into HTML.
  const escapeHtml = (value: string): string => {

    return value.replace(/[&<>"']/g, (char: string) => {

      const entities: { [key: string]: string } = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
      };

      return entities[char];
    });
  };

  // Resolve SharePoint Hyperlink column values.
  const getMenuUrl = (item: IMenuItem): string => {

    const redirectUrl = item.RedirectURL;

    if (redirectUrl && typeof redirectUrl === 'object') {
      return redirectUrl.Url || '#';
    }

    return (redirectUrl as unknown as string) || '#';
  };

  // Identify top-level menu items.
  const parentItems = this._menuItems.filter(
    (item: IMenuItem) =>
      !item.ParentMenu || !item.ParentMenu.Id
  );

  // Get direct children of a menu item.
  const getChildItems = (parentId: number): IMenuItem[] => {

    return this._menuItems.filter(
      (item: IMenuItem) =>
        !!item.ParentMenu &&
        Number(item.ParentMenu.Id) === Number(parentId)
    );
  };

  // Recursively render menu items and nested dropdowns.
  const renderMenuItem = (
    item: IMenuItem,
    ancestorIds: number[] = [],
    isNested: boolean = false
  ): string => {

    // Prevent infinite recursion if the list contains a circular parent reference.
    if (ancestorIds.indexOf(item.Id) !== -1) {
      return '';
    }

    const title = escapeHtml(getMenuTitle(item));
    const url = escapeHtml(getMenuUrl(item));

    const childItems = getChildItems(item.Id);

    // Render a normal link when there are no children.
    if (childItems.length === 0) {

      return `
        <a
          href="${url}"
          class="${isNested ? styles.dropdownItem : styles.navLink}">
          ${title}
        </a>
      `;
    }

    // Render each child, including any further nested levels.
    const nextAncestorIds = [...ancestorIds, item.Id];

    const childMenuHtml = childItems
      .map((child: IMenuItem) =>
        renderMenuItem(child, nextAncestorIds, true)
      )
      .join('');

    return `
      <div class="${styles.dropdown}">

        <a
          href="${url}"
          class="${isNested ? styles.dropdownItem : styles.navLink}"
          aria-haspopup="true">
          ${title}
          
        </a>

        <div class="${styles.dropdownMenu}">
          ${childMenuHtml}
        </div>

      </div>
    `;
  };

  // Generate the complete navigation menu.
  const menuHtml = parentItems
    .map((item: IMenuItem) => renderMenuItem(item))
    .join('');

  this._top.domElement.innerHTML = `

    <div id="bpcl-custom-navbar" class="${styles.mainNavbar}">

      <div class="${styles.navbarContainer}">

        <div class="${styles.logoSection}">
          <img
            src="${Logo}"
            alt="MRIonnect Logo"
            class="${styles.logoImg}" />
        </div>

        <button
          id="bpclMenuBtn"
          class="${styles.menuBtn}"
          aria-label="Toggle Menu"
          aria-expanded="false">
          ☰
        </button>

        <div id="bpclMenu" class="${styles.navMenu}">

          <div class="${styles.navLinks}">
            ${menuHtml}
          </div>

          <div class="${styles.languageWrapper}">

            <span
              id="mrCurrentDateTime"
              class="${styles.timestamp}">
              ${this._getCurrentDateTime()}
            </span>

            <select
              id="mrLanguageSelector"
              class="${styles.languageSelect}">

              <option value="English"
                ${selectedLanguage === 'English' ? 'selected' : ''}>
                English
              </option>

              <option value="Hindi"
                ${selectedLanguage === 'Hindi' ? 'selected' : ''}>
                Hindi
              </option>

              <option value="Marathi"
                ${selectedLanguage === 'Marathi' ? 'selected' : ''}>
                Marathi
              </option>

            </select>

          </div>

        </div>

      </div>

    </div>
  `;

  // Mobile menu toggle.
  const menuBtn = document.getElementById(
    'bpclMenuBtn'
  ) as HTMLButtonElement | null;

  const menu = document.getElementById('bpclMenu');

  if (menuBtn && menu) {

    menuBtn.addEventListener('click', () => {

      const isOpen =
        menu.classList.toggle('bpcl-mobile-open');

      menuBtn.textContent = isOpen ? '✕' : '☰';

      menuBtn.setAttribute(
        'aria-expanded',
        String(isOpen)
      );
    });
  }

  // Language selector.
  const languageSelector = document.getElementById(
    'mrLanguageSelector'
  ) as HTMLSelectElement | null;

  if (languageSelector) {

    languageSelector.addEventListener(
      'change',
      async (e: Event) => {

        const selectedValue =
          (e.target as HTMLSelectElement).value;

        localStorage.setItem(
          'MRLanguage',
          selectedValue
        );

        if (this._top) {
          this._top.dispose();
          this._top = undefined;
        }

        await this._renderTop();

        window.dispatchEvent(
          new CustomEvent('languageChanged')
        );
      }
    );
  }
}


  private _renderBottom(): void {

    if (document.getElementById('bpcl-footer')) {
      return;
    }

    const pageCanvas = document.querySelector(
      '#spPageCanvasContent'
    );

    if (!pageCanvas) {
      return;
    }

    const year = new Date().getFullYear();

    const footer = document.createElement('div');

    footer.id = 'bpcl-footer';

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

    pageCanvas.appendChild(footer);
  }
}
