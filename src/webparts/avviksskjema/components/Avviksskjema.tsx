import * as React from 'react';
import { IAvviksskjemaProps } from './IAvviksskjemaProps';
import { IAvviksskjemaState, DefatultState } from './IAvviksskjemaState';
import {
  ChoiceGroup,
  ComboBox,
  DatePicker,
  DayOfWeek,
  DefaultButton,
  IChoiceGroupOption,
  IComboBoxOption,
  IDatePickerProps,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  SelectableOptionMenuItemType,
  Spinner,
  Stack,
  TextField
} from '@fluentui/react';
import { dateAdd, PnPClientStorage } from '@pnp/common';
import {
  DateTimePicker,
  TimeConvention,
  TimeDisplayControlType,
  MinutesIncrement,
  IDateTimePickerProps,
} from '@pnp/spfx-controls-react';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http'; 

const PnpStorage = new PnPClientStorage();
const PnpStorageKey = 'Avviksskjema';

export interface ISalesforceErrorRespone {
  errorCode: string;
  message: string;
}

export default class Avviksskjema extends React.Component<IAvviksskjemaProps, IAvviksskjemaState> {

  public constructor(props: IAvviksskjemaProps) {
    super(props);
    this.state = DefatultState;
    this._loadState();
  }

  public componentDidUpdate() {
    this._saveState();
  }
  
  public async componentDidCatch(error, info) {
    this.setState({hasError: true, sending: false});
    console.error(error);
    console.info(info);
  }
  
  public render(): React.ReactElement<IAvviksskjemaProps> {
    
    const disabledIfCategoryIsEmpty = {
      disabled: this.state.category === '',
    };

    const shortTextFieldProps = {
      onGetErrorMessage: (value: string): string => this._getErrorMessageTextLength(value, 200),
      ...disabledIfCategoryIsEmpty,
    };

    const longTextFieldProps = {
      multiline: true,
      autoAdjustHeight: true,
      onGetErrorMessage: (value: string): string => this._getErrorMessageTextLength(value, 32768),
      ...disabledIfCategoryIsEmpty,
    };

    const choiceGroupProps = {
      ...disabledIfCategoryIsEmpty,
    };

    const dateLocalizationProps: IDatePickerProps = {
      strings: {
        goToToday: 'Gå til i dag',
        days: ['Søndag', 'Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag', 'Lørdag'],
        shortDays: ['Søn', 'Man', 'Tir', 'Ons', 'Tor', 'Fre', 'Lør'],
        months: ['Januar', 'Februar', 'Mars', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Desember'],
        shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Des'],
        dateLabel: 'Dato',
        timeLabel: 'Klokkeslett',
        timeSeparator: ':',
        textErrorMessage: 'Ikke skriv tekst her',
      },
      formatDate: (date?: Date) => date && date.toLocaleDateString(),
      firstDayOfWeek: DayOfWeek.Monday,
      ...disabledIfCategoryIsEmpty,
    } as IDatePickerProps;

    const dateTimeLocalizationProps: IDateTimePickerProps = {
      ...dateLocalizationProps as unknown as IDateTimePickerProps,
      timeConvention: TimeConvention.Hours24,
      timeDisplayControlType: TimeDisplayControlType.Dropdown,
      minutesIncrementStep: 10 as MinutesIncrement,
    };

    const categoryOptions: IComboBoxOption[] = [
      { key: 'Sikkerhetsavvik', text: 'Sikkerhetsavvik', itemType: SelectableOptionMenuItemType.Header },
      { key: 'Personopplysninger på avveie', text: 'Personopplysninger på avveie' },
      { key: 'Brudd på policy/retningslinje', text: 'Brudd på policy/retningslinje' },
      { key: 'Mangel på policy/retningslinje', text: 'Mangel på policy/retningslinje' },
      { key: 'Security Exception/fravik', text: 'Security Exception/fravik' },
      { key: 'Andre hendelser', text: 'Andre hendelser' },
      { key: 'HMS', text: 'HMS', itemType: SelectableOptionMenuItemType.Header },
      { key: 'Vold og trusler', text: 'Vold og trusler' },
      { key: 'HMS avvik', text: 'HMS avvik' },
      { key: 'HMS forbedringsforslag', text: 'HMS forbedringsforslag' },
      { key: 'Personvern', text: 'Personvern', itemType: SelectableOptionMenuItemType.Header },
      { key: 'Manglende behandlingsgrunnlag', text: 'Manglende behandlingsgrunnlag' },
      { key: 'Manglende ivaretagelse av innsynsretten', text: 'Manglende ivaretagelse av innsynsretten' },
      { key: 'Mangler databehandleravtale', text: 'Mangler databehandleravtale' },
      { key: 'Ugyldig samtykke', text: 'Ugyldig samtykke' },
      { key: 'Ikke fastsatt lagringstider', text: 'Ikke fastsatt lagringstider' },
      { key: 'Behandlingen er ikke registrert i Behandlingskatalogen', text: 'Behandlingen er ikke registrert i Behandlingskatalogen' },
      { key: 'Ikke gjennomført personvernkonsekvensvurdering (PVK)', text: 'Ikke gjennomført personvernkonsekvensvurdering (PVK)' },
    ];
  
    const priorityOptions: IChoiceGroupOption[] = [
      { key: 'Lav', text: 'Lav' },
      { key: 'Middels', text: 'Middels' },
      { key: 'Høy', text: 'Høy'},
    ];
  
    const incidentMainCauseOptions: IChoiceGroupOption[] = [
      { key: 'Brudd på rutiner', text: 'Brudd på rutiner' },
      { key: 'Manglende rutiner', text: 'Manglende rutiner' },
      { key: 'Menneskelig svikt', text: 'Menneskelig svikt'},
      { key: 'Teknisk svikt', text: 'Teknisk svikt'},
      { key: 'Annet', text: 'Annet'},
    ];
  
    const relationsForPeopleInvolvedOptions: IChoiceGroupOption[] = [
      { key: 'Ansatt/Innleid', text: 'Ansatt/Innleid' },
      { key: 'NAV-bruker', text: 'NAV-bruker' },
      { key: 'Annet', text: 'Annet'},
    ];
  
    const isRelatedToSecurityLawOptions: IChoiceGroupOption[] = [
      { key: 'Ja', text: 'Ja' },
      { key: 'Nei', text: 'Nei' },
      { key: 'Vet ikke', text: 'Vet ikke'},
    ];
      
    return (<form onSubmit={this.sendForm}>
      <Stack tokens={{ childrenGap: 20}}>
        <ComboBox 
          label='Hvilken kategori gjelder avviket?'
          options={categoryOptions}
          selectedKey={this.state.category}
          onChange={(_, opt) => this.setState({category: opt.key as string})}
        />
        <DatePicker 
          label='Når skjedde/startet hendelsen?'
          onSelectDate={val => this.setState({incidentDate: val})}
          value={this.state.incidentDate}
          {...dateLocalizationProps}
        />
        <TextField 
          label='Hvor skjedde hendelsen?'
          description='Enhet / Geografisk lokasjon'
          value={this.state.incidentLocation}
          onChange={(_, val) => this.setState({incidentLocation: val})}
          {...shortTextFieldProps}
        />
        <TextField 
          label='Beskriv hendelsen'
          value={this.state.incidentDescription}
          onChange={(_, val) => this.setState({incidentDescription: val})}
          {...longTextFieldProps}
        />
        <TextField
          label='Hvilke konsekvenser hadde hendelsen?'
          value={this.state.incidentConsecquences}
          onChange={(_, val) => this.setState({incidentConsecquences: val})}
          {...longTextFieldProps}
        />
        <TextField
          label='Hva er årsaken til hendelsen?'
          value={this.state.incidentCause}
          onChange={(_, val) => this.setState({incidentCause: val})}
          {...longTextFieldProps}
        />
        <TextField
          label='Forslag til tiltak:'
          value={this.state.suggestedActions}
          onChange={(_, val) => this.setState({suggestedActions: val})}
          {...longTextFieldProps}
        />
        <ChoiceGroup 
          label='Alvorlighetsgrad'
          options={priorityOptions}
          selectedKey={this.state.priority}
          onChange={(_, val) => this.setState({priority: val.key})}
          {...choiceGroupProps}
        />
        {this.state.category === categoryOptions[1].key && <>
        <TextField
          label='Hvem er de berørte?'
          description='Oppgi navn og personnummer. Ett per linje.'
          value={this.state.peopleInvolved}
          onChange={(_, val) => this.setState({peopleInvolved: val})}
          {...longTextFieldProps}
        />
        <DatePicker 
          label='Hvor lenge varte hendelsen?'
          value={this.state.incidentToDate}
          onSelectDate={val => this.setState({incidentToDate: val})}
          {...dateLocalizationProps}
          />
        <DateTimePicker 
          label='Når ble avviket oppdaget?'
          value={this.state.incidentFoundDateTime}
          onChange={val => this.setState({incidentFoundDateTime: val})}
          {...dateTimeLocalizationProps}
        />
        <ChoiceGroup 
          label='Hovedårsak'
          options={incidentMainCauseOptions}
          selectedKey={this.state.incidentMainCause}
          onChange={(_, val) => this.setState({incidentMainCause: val.key})}
          {...choiceGroupProps}
        />
        <ChoiceGroup 
          label='Hvilken relasjon har virksomheten til de personene som er berørt av avviket?'
          options={relationsForPeopleInvolvedOptions}
          selectedKey={this.state.relationsForPeopleInvolved}
          onChange={(_, val) => this.setState({relationsForPeopleInvolved: val.key})}
          {...choiceGroupProps}
        />
        { this.state.relationsForPeopleInvolved &&
          this.state.relationsForPeopleInvolved === relationsForPeopleInvolvedOptions[2].key &&
        <>
          <TextField
            label='Du valgte «annet». Vennligst spesifiser:'
            value={this.state.relationsForPeopleInvolvedOther}
            onChange={(_, val) => this.setState({relationsForPeopleInvolvedOther: val})}
            {...shortTextFieldProps}
          />
        </>}
        </>}
        {this.state.category === categoryOptions[2].key && <>
        <ChoiceGroup
          label='Er hendelsen relatert til sikkerhetsloven?'
          options={isRelatedToSecurityLawOptions}
          selectedKey={this.state.isRelatedToSecurityLaw}
          onChange={(_, val) => this.setState({isRelatedToSecurityLaw: val.key})}
          {...choiceGroupProps}
        />
        <DatePicker 
          label='Hvor lenge varte hendelsen?'
          value={this.state.incidentToDate}
          onSelectDate={val => this.setState({incidentToDate: val})}
          {...dateLocalizationProps}
          />
        <TextField
          label='Hvilken enhet er berørt?'
          value={this.state.involvedUnit}
          onChange={(_, val) => this.setState({involvedUnit: val})}
          {...shortTextFieldProps}
        />
        </>}
        {this.state.category === categoryOptions[3].key && <>
        <TextField
          label='Hvem eller hva er berørt?'
          description='Enhet, system eller enkeltperson'
          value={this.state.involved}
          onChange={(_, val) => this.setState({involved: val})}
          {...longTextFieldProps}
        />
        <TextField
          label='Hvilken policy mangler?'
          value={this.state.missingPolicy}
          onChange={(_, val) => this.setState({missingPolicy: val})}
          {...longTextFieldProps}
        />
        <TextField
          label='Hva ønskes oppnådd? Hva er behovet?'
          value={this.state.resultNeeds}
          onChange={(_, val) => this.setState({resultNeeds: val})}
          {...longTextFieldProps}
        />
        <TextField
          label='Forslag til løsning:'
          value={this.state.suggestedResolution}
          onChange={(_, val) => this.setState({suggestedResolution: val})}
          {...longTextFieldProps}
        />
        </>}
        {this.state.category === categoryOptions[4].key && <>
        <TextField
          label='Hvem eller hva er berørt? (Enhet, system eller enkeltperson)'
          value={this.state.involved}
          onChange={(_, val) => this.setState({involved: val})}
          {...longTextFieldProps}
        />
        <TextField
          label='Hvilken policy fravikes?'
          value={this.state.missingPolicy}
          onChange={(_, val) => this.setState({missingPolicy: val})}
          {...longTextFieldProps}
        />
        <TextField
          label='Hva er begrunnelsen? Hvorfor er det fravik?'
          value={this.state.resultNeeds}
          onChange={(_, val) => this.setState({resultNeeds: val})}
          {...longTextFieldProps}
        />
        <TextField
          label='Beskriv kompenserende tiltak:'
          value={this.state.suggestedResolution}
          onChange={(_, val) => this.setState({suggestedResolution: val})}
          {...longTextFieldProps}
        />
        </>}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton 
            text='Send'
            type='submit'
          />
          {this.state.sending && <Spinner label='Sender…' ariaLive="assertive" labelPosition="right" />}
        </Stack>
        {this.state.hasError && <>
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={()=>this.setState({hasError: false})}
          isMultiline={true}
        >
          {this.state.errorMessage && <>
            Det skjedde en feil i innsendingen. <br />
            <br />
            {this.state.errorCode && <><strong>Feilkode:</strong> {this.state.errorCode}<br /></>}
            <strong>Feilmelding:</strong> {this.state.errorMessage}
          </>}
          {this.state.hasError && !this.state.errorMessage && <>
            Klarte ikke å få kontakt med SalesForce. Se nettleserkonsollen for detaljer.
          </>}
        </MessageBar>
        </>}
        {this.state.responseID && <>
        <MessageBar
          messageBarType={MessageBarType.success}
          onDismiss={()=>this.setState({responseID: undefined})}
        >
          Vellykket innsending. <strong>Saksnummer:</strong> {this.state.responseID}.
        </MessageBar>
        </>}
        <Stack tokens={{ childrenGap: 10 }}>
          <TextField
            label='Skjemadata'
            value={JSON.stringify({...this._getFormFields(), reporterID: this.props.context.pageContext.user.loginName}, undefined, 2)}
            readOnly multiline autoAdjustHeight
          />
          <DefaultButton
            text='Slett skjemadata fra nettleseren'
            onClick={this._deleteState}
          />
        </Stack>
      </Stack>
    </form>);
  }

  private _getErrorMessageTextLength = (value: string, limit: number): string => {
    return value.length < limit ? '' : `Antall tegn må være mindre enn ${limit}. Antall tegn er nå ${value.length}.`;
  }

  private _loadState = async () => {
    await PnpStorage.session.deleteExpired();
    const storedState = await PnpStorage.session.get(PnpStorageKey);
    storedState.incidentDate = storedState.incidentDate && new Date(storedState.incidentDate);
    storedState.incidentToDate = storedState.incidentToDate && new Date(storedState.incidentToDate);
    storedState.incidentFoundDateTime = storedState.incidentFoundDateTime && new Date(storedState.incidentFoundDateTime);
    this.setState(storedState);
  }

  private _saveState = () => {
    PnpStorage.session.put(PnpStorageKey, this._getFormFields(), dateAdd(new Date(), 'day', 1));
  }

  private _deleteState = () => {
    PnpStorage.session.delete(PnpStorageKey);
    this.setState(DefatultState);
  }

  protected sendForm = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!this.props.salesforceUrl || !this.props.salesforceToken) {
      this.setState({
        hasError: true,
        errorMessage: 'Mangler url eller token til Salesforce-API. Dette må legges inn i nettdelens innstillinger.',
      });
      return;
    }
    this.setState({sending: true});
    const body = JSON.stringify(this._getFormFields());
    const headers: Headers = new Headers({
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${this.props.salesforceToken.trim()}`,
      'x-prettyprint': '1',
    });
    const httpClientOptions: IHttpClientOptions = {body, headers};
    try {
      const response: HttpClientResponse = await this.props.context.httpClient.post(
        this.props.salesforceUrl.trim(),
        HttpClient.configurations.v1,
        httpClientOptions,
      );
      const json: string | ISalesforceErrorRespone[] = await response.json();
      if (typeof json === "string") {
        // success!
        this.setState({hasError: false, responseID: json});
      } else {
        // error
        this.setState({hasError: true, errorCode: json[0].errorCode, errorMessage: json[0].message});
      }
    } catch (e) {
      this.setState({
        hasError: true,
        errorMessage: `${e}. Sjekk nettleserkonsollen for mer informasjon.`,
      });
      console.error(e);
    } finally {
      this.setState({sending: false});
    }
  }

  private _getFormFields = () => {
    const fields = {...this.state}; // clone
    [
      'hasError',
      'responseID',
      'errorCode',
      'errorMessage',
      'sending',
    ].forEach(k => delete fields[k]);
    return fields;
  }

}
