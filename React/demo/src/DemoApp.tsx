// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useEffect } from 'react';
import { models, Report, Embed, service, Page } from 'powerbi-client';
import { IHttpPostMessageResponse } from 'http-post-message';
import { PowerBIEmbed } from 'powerbi-client-react';
import 'powerbi-report-authoring';

import { sampleReportUrl } from './public/constants';
import './DemoApp.css';

// Root Component to demonstrate usage of embedded component
function DemoApp (): JSX.Element {

	// PowerBI Report object (to be received via callback)
	const [report, setReport] = useState<Report>();

	// Track Report embedding status
	const [isEmbedded, setIsEmbedded] = useState<boolean>(false);

    // Overall status message of embedding
	const [displayMessage, setMessage] = useState(`The report is bootstrapped. Click the Embed Report button to set the access token`);

	// CSS Class to be passed to the embedded component
	const reportClass = 'report-container';

	// Pass the basic embed configurations to the embedded component to bootstrap the report on first load
    // Values for properties like embedUrl, accessToken and settings will be set on click of button
	// const [sampleReportConfig, setReportConfig] = useState<models.IReportEmbedConfiguration>({
	// 	type: 'report',
	// 	embedUrl: undefined,
	// 	tokenType: models.TokenType.Embed,
	// 	accessToken: undefined,
	// 	settings: undefined,
	// });
		const [sampleReportConfig, setReportConfig] = useState<models.IReportEmbedConfiguration>({
		type: 'report',
		embedUrl: 'https://app.powerbi.com/reportEmbed',
		tokenType: models.TokenType.Embed,
		accessToken: 'H4sIAAAAAAAEAB2TN66EVgAA7_JbLJGTJRfk_Mhh6RaWnPOC5bv72_1UM5q_f5z33U_vz8-fPx2ZPqWIiGUXBbbRhQGtloLnvQn9ImSVR0rPTm-9Nun11Blnj2DUs0Lu3BlokAPGCXImj5s6JZB4FMIixszat7wA-bhwCH1sGahhodN23TfZMWl56YLeH48waaOSMFMrApiy12FFO2KT4FohV1u86W2COM-afIQyCDNLGl4z7H9OQ3Pt0szwTVyn2lWUal90OHMJgYNq_L6rj_xoCvcFTJj17P5w6cRyUtD0juQaYrtFtZkZzLqtEIPa1v1ZdtOgMUebdFNw7JaKFCEtgdm6rFwNclOBbUcFv1QJgiZDfQdNIhd6O7fThgCsFILtFWLn8WDga90-0pLC4lAKvkEK6SYF83F1xG9r0qE7u7NNowSFmW7OO3mOCcf5bh8CzPKbopllyEkdBtN4Op1Rr7L3pUFzDAnbVcMVQZUKhoKr3IOfRlRLfMuN2l6xIze64j7nW6bFVbEbbhp1z4SJ2OME2MklBlWrDo0roiik8oDCa66nNdAr5BslMQLs4UYH7uAQ22fR6g1T5wO2wpNpqIfCDXe0oYkWbX4aRyK15k0v3_IggI_N8uUALM3pHMDgQjR6yfwOxXqzNPuRw046YW-rfz8NDbK9hU8-q1LrroprohoySrEebWeHJEBNTZCVk3JbpzQXfOG3imZjBIRNfS33dnczbDmO09FPCrQFpZWPcR3XqgmJdrjYtNGE7irqea2P7Bb8WVtqil28Pb5PFoKxboCX1GDL8ItIGYcZDuhG1NeD-ssamZxspXdQtHzuId6lWYa7gnyuSHY7EpxkyaUdtbpyLsCFya1QHtYJNQM8osDqhDax6H6OgFGwfKcKaLdCzf7540dY73mfjOL-3em3g47wUldW3eAgLhHlEL0hFtBVWQ1YnkCjhoZxzbaorbNvr3ci4J281LiNJVWRqs6mtlPUDhQ93RZCNjm9QeY5daFzRfGY9GVL5O-LVUwC8ao1WllGbXxP0brh9vhLnyn7zSpvJhNGlR4ELcNgHrHXXB7ioBAFxI1Ix1o_FDaGI3kHT9dvjt2RzsB631xkRu-Dh6V8QZXBUTns8-89uFytozCkfNXqpIZ75R11r6TnwZTHgDvGN9rYbUzGfnR8rftKuL994-jgYxvPmeSCxcaVzmDbOQXnlIit1rGeAg_SxaiRcGJb8DLjrGXvUEkXb-uzVqI0HJK0E_DFH6K78eA-3b_--k_zPdfFqkW_lp3VO1vBe61a8Dx4a0R7WfPc_5TfVON7P9biF-MtyuwVmox3GBh3AGdGjLMUN4aupYblwyHI0Hc0s0xqPBHqJNgXvG5D1X3RkqP8tA6FKExwvH4G43DmLT_0_kUlNqS7KA4HN8cdDb-IJFt_sc6PV29Z1VNJ8bua0zLxLgvkoXS4X04Z5_CwqHkbNV-Zhx4aoRc-nONqaxkwy5kzK2PDPVSQBxRKeRSqP_DGx3QWw9TryqH7FU4BhgeCH-nsb6IFGCaXExtskLZAf7C90m9mASlXN4nwXfn6eBPY8dUKuGYDcJaTvHSODckaLZYKTjmK-K5AAKAuOtJs18uqxe2q9RRLu7GMNA1rCeoN5ObLy4DO8P5pxhoR4swOQuv5L8Y__wKDhdifQgYAAA.eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly9XQUJJLVdFU1QtVVMtcmVkaXJlY3QuYW5hbHlzaXMud2luZG93cy5uZXQiLCJleHAiOjE2ODU5ODY2NTAsImFsbG93QWNjZXNzT3ZlclB1YmxpY0ludGVybmV0Ijp0cnVlfQ',
		// settings: undefined,
		id: 'f8188497-ce38-480c-b5c0-239f52b42be1'
	});

	/**
	 * Map of event handlers to be applied to the embedded report
	 * Update event handlers for the report by redefining the map using the setEventHandlersMap function
	 * Set event handler to null if event needs to be removed
	 * More events can be provided from here
	 * https://docs.microsoft.com/en-us/javascript/api/overview/powerbi/handle-events#report-events
	 */
	const[eventHandlersMap, setEventHandlersMap] = useState<Map<string, (event?: service.ICustomEvent<any>, embeddedEntity?: Embed) => void | null>>(new Map([
		['loaded', () => console.log('Report has loaded')],
		['rendered', () => console.log('Report has rendered')],
		['error', (event?: service.ICustomEvent<any>) => {
				if (event) {
					console.error(event.detail);
				}
			},
		],
		['visualClicked', () => console.log('visual clicked')],
		['pageChanged', (event) => console.log(event)],
	]));

	useEffect(() => {
		if (report) {
			report.setComponentTitle('Embedded Report');
		}
	}, [report]);

    /**
     * Embeds report
     *
     * @returns Promise<void>
     */
	const embedReport = async (): Promise<void> => {
		console.log('Embed Report clicked');

		// Get the embed config from the service
		const reportConfigResponse = await fetch(sampleReportUrl);

		if (reportConfigResponse === null) {
			return;
		}

		if (!reportConfigResponse?.ok) {
			console.error(`Failed to fetch config for report. Status: ${ reportConfigResponse.status } ${ reportConfigResponse.statusText }`);
			return;
		}

		const reportConfig = await reportConfigResponse.json();

		// Update the reportConfig to embed the PowerBI report
		setReportConfig({
			...sampleReportConfig,
			embedUrl: reportConfig.EmbedUrl,
			accessToken: reportConfig.EmbedToken.Token
		});
		setIsEmbedded(true);

		// Update the display message
		setMessage('Use the buttons above to interact with the report using Power BI Client APIs.');
	};

    /**
     * Hide Filter Pane
     *
     * @returns Promise<IHttpPostMessageResponse<void> | undefined>
     */
	const hideFilterPane = async (): Promise<IHttpPostMessageResponse<void> | undefined>  => {
		// Check if report is available or not
		if (!report) {
			setDisplayMessageAndConsole('Report not available');
			return;
		}

		// New settings to hide filter pane
		const settings = {
			panes: {
				filters: {
					expanded: false,
					visible: false,
				},
			},
		};

		try {
			const response: IHttpPostMessageResponse<void> = await report.updateSettings(settings);

			// Update display message
			setDisplayMessageAndConsole('Filter pane is hidden.');
			return response;
		} catch (error) {
			console.error(error);
			return;
		}
	};

    /**
     * Set data selected event
     *
     * @returns void
     */
	const setDataSelectedEvent = () => {
		setEventHandlersMap(new Map<string, (event?: service.ICustomEvent<any>, embeddedEntity?: Embed) => void | null> ([
			...eventHandlersMap,
			['dataSelected', (event) => console.log(event)],
		]));

		setMessage('Data Selected event set successfully. Select data to see event in console.');
	}

    /**
     * Change visual type
     *
     * @returns Promise<void>
     */
	const changeVisualType = async (): Promise<void> => {
		// Check if report is available or not
		if (!report) {
			setDisplayMessageAndConsole('Report not available');
			return;
		}

		// Get active page of the report
		const activePage: Page | undefined = await report.getActivePage();

		if (!activePage) {
			setMessage('No Active page found');
			return;
		}

		try {
			// Change the visual type using powerbi-report-authoring
			// For more information: https://docs.microsoft.com/en-us/javascript/api/overview/powerbi/report-authoring-overview
			const visual = await activePage.getVisualByName('VisualContainer6');

			const response = await visual.changeType('lineChart');

			setDisplayMessageAndConsole(`The ${visual.type} was updated to lineChart.`);

			return response;
		}
		catch (error) {
			if (error === 'PowerBIEntityNotFound') {
				console.log('No Visual found with that name');
			} else {
				console.log(error);
			}
		}
	};

	/**
     * Set display message and log it in the console
     *
     * @returns void
     */
	const setDisplayMessageAndConsole = (message: string): void => {
		setMessage(message);
		console.log(message);
	}

	const controlButtons =
		isEmbedded ?
		<>
			<button onClick = { changeVisualType }>
				Change visual type</button>

			<button onClick = { hideFilterPane }>
				Hide filter pane</button>

			<button onClick = { setDataSelectedEvent }>
				Set event</button>

			<label className = "display-message">
				{ displayMessage }
			</label>
		</>
		:
		<>
			<label className = "display-message position">
				{ displayMessage }
			</label>

			<button onClick = { embedReport } className = "embed-report">
				Embed Report</button>
		</>;

	const header =
		<div className = "header">Power BI Embedded React Component Demo</div>;

	const reportComponent =
		<PowerBIEmbed
			embedConfig = { sampleReportConfig }
			eventHandlers = { eventHandlersMap }
			cssClassName = { reportClass }
			getEmbeddedComponent = { (embedObject: Embed) => {
				console.log(`Embedded object of type "${ embedObject.embedtype }" received`);
				setReport(embedObject as Report);
			} }
		/>;

	const footer =
		<div className = "footer">
			<p>This demo is powered by Power BI Embedded Analytics</p>
			<label className = "separator-pipe">|</label>
			<img title = "Power-BI" alt = "PowerBI_Icon" className = "footer-icon" src = "./assets/PowerBI_Icon.png" />
			<p>Explore our<a href = "https://aka.ms/pbijs/" target = "_blank" rel = "noreferrer noopener">Playground</a></p>
			<label className = "separator-pipe">|</label>
			<img title = "GitHub" alt = "GitHub_Icon" className = "footer-icon" src = "./assets/GitHub_Icon.png" />
			<p>Find the<a href = "https://github.com/microsoft/PowerBI-client-react" target = "_blank" rel = "noreferrer noopener">source code</a></p>
		</div>;

	return (
		<div className = "container">
			{ header }

			<div className = "controls">
				{ controlButtons }

				{ isEmbedded ? reportComponent : null }
			</div>

			{ footer }
		</div>
	);
}

export default DemoApp;