import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import { Dialog } from "@microsoft/sp-dialog";

import * as strings from "AppExtentionApplicationCustomizerStrings";

const LOG_SOURCE: string = "AppExtentionApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppExtentionApplicationCustomizerProperties {
	// This is an example; replace with your own property
	titleName: string;
	favicon: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppExtentionApplicationCustomizer extends BaseApplicationCustomizer<IAppExtentionApplicationCustomizerProperties> {
	public onInit(): Promise<void> {
		Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

		let title: string = this.properties.titleName;
		//set page title
		if (!title) {
			title = "(No properties were provided.)";
		} else {
			document.title = title;
		}

		//set page favicon
		const faviconUrl: string = this.properties.favicon;
		if (!faviconUrl) {
			Log.info(LOG_SOURCE, "Favicon is missing!");
		} else {
			const link: HTMLElement = (document.querySelector("link[rel*='icon']") as HTMLElement) || (document.createElement("link") as HTMLElement);
			link.setAttribute("type", "image/x-icon");
			link.setAttribute("rel", "shortcut icon");
			link.setAttribute("href", faviconUrl);
			document.getElementsByTagName("head")[0].appendChild(link);
		}

		// Dialog.prompt(strings.Title).catch(() => {});
		// Dialog.alert(strings.Title).catch(() => {
		/* handle error */
		// });

		return Promise.resolve();
	}
}
