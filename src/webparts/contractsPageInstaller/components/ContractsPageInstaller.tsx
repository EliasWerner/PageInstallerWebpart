import * as React from 'react';
import styles from './ContractsPageInstaller.module.scss';
import { IContractsPageInstallerProps } from './IContractsPageInstallerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { sp, ClientSideText, ClientSideWebpart } from '@pnp/pnpjs';

const contractsWebPartId = '66B071BF-ABDC-42A4-8A47-154A53878240';
const title = 'Active contracts page';

interface IInstallerState {
	isCreateingPage: boolean;
}

export default class ContractsPageInstaller extends React.Component<IContractsPageInstallerProps, IInstallerState> {
	private textFieldInputText = '';

	constructor(props: IContractsPageInstallerProps) {
		super(props);

		this.state = {
			isCreateingPage: false
		};
	}
	public render(): React.ReactElement<IContractsPageInstallerProps> {
		return (
			<div className={styles.contractsPageInstaller}>
				<div className={styles.container}>
					<div className={styles.row}>
						<div className={styles.column}>
							<span className={styles.title}>Welcome</span>
							<p className={styles.subTitle}>
								This app will help you to create page with list of current active contracts. Please, specify the name of the new page.
							</p>
							<div className={styles.inline}>
								<TextField onChange={this.onTextFieldChange} placeholder="Page name" style={{ width: 500 }} />
								<button className={styles.button} title="Install" onClick={this.onInstallClick} />
							</div>
						</div>
					</div>
					{this.state.isCreateingPage ? (
						<div className={styles.row}>
							<Spinner size={SpinnerSize.large} />
						</div>
					) : null}
				</div>
			</div>
		);
	}

	private onInstallClick = async () => {
		if (!this.textFieldInputText.length) {
			alert('Please, specify the name of the page.');
			return;
		}

		await this.checkIfPageExists(this.textFieldInputText);

		this.setState({ isCreateingPage: true });

		const page = await sp.web.addClientSidePage(this.textFieldInputText, title);
		page.addSection();
		const section = page.addSection();
		section.addColumn(6);

		const partDefs = await sp.web.getClientSideWebParts();
		const res = partDefs.filter(t => {
			return t.Id.indexOf(contractsWebPartId) !== -1;
		})[0];

		const part = ClientSideWebpart.fromComponentDef(res);
		page.addSection().addControl(part);
		await page.save();

		this.setState({ isCreateingPage: false });
	};

	private onTextFieldChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
		this.textFieldInputText = text;
	};

	private checkIfPageExists = async (pageName: string) => {
		const sitePages = await sp.web.getList('SitePages').items.get();
		console.log(sitePages);
	};
}
