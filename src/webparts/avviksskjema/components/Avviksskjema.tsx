import * as React from 'react';
import { IAvviksskjemaProps } from './IAvviksskjemaProps';
import { IAvviksskjemaState, DefatultState } from './IAvviksskjemaState';
import * as strings from 'AvviksskjemaWebPartStrings';
import {
  Checkbox,
  ChoiceGroup,
  DatePicker,
  DayOfWeek,
  DefaultButton,
  IChoiceGroupOption,
  IChoiceGroupProps,
  IDatePickerProps,
  IDatePickerStrings,
  ITextFieldProps,
  ITextFieldStyleProps,
  ITextFieldStyles,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  Stack,
  TextField,
  Toggle
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
    
    const getTextFieldStyles = (props: ITextFieldStyleProps): Partial<ITextFieldStyles> => {
      const {required} = props;
      return {
        fieldGroup: [],
        subComponentStyles: {
          label: {
            root: {
              // fontSize: '16px'
            }
          },
        }
      };
    };

    const shortTextFieldProps: ITextFieldProps = {
      onGetErrorMessage: (value: string): string => this._getErrorMessageTextLength(value, 200),
      styles: getTextFieldStyles,
    };

    const longTextFieldProps: ITextFieldProps = {
      multiline: true,
      autoAdjustHeight: true,
      // onGetErrorMessage: (value: string): string => this._getErrorMessageTextLength(value, 32768),
      styles: getTextFieldStyles,
    };

    const choiceGroupProps: IChoiceGroupProps = {
      styles: getTextFieldStyles,
    };

    const dateLocalizationProps: IDatePickerProps = {
      strings: strings.DateStrings as unknown as IDatePickerStrings,
      formatDate: (date?: Date) => date && date.toLocaleDateString(),
      firstDayOfWeek: DayOfWeek.Monday,
      styles: getTextFieldStyles,
    } as IDatePickerProps;

    const dateTimeLocalizationProps: IDateTimePickerProps = {
      ...dateLocalizationProps as unknown as IDateTimePickerProps,
      timeConvention: TimeConvention.Hours24,
      timeDisplayControlType: TimeDisplayControlType.Dropdown,
      minutesIncrementStep: 10 as MinutesIncrement,
    };

    const options = (labels: string[]): IChoiceGroupOption[] => {
      return labels.map(label => ({ key: label, text: label }));
    };

    const categoryOptions: IChoiceGroupOption[] = options([
      strings.IncidentCategoryPrivacy,
      strings.IncidentCategorySecurity,
      strings.IncidentCategoryHSE,
      strings.IncidentCategoryOther,
    ]);
  
    const incidentMainCauseOptions: IChoiceGroupOption[] = options([
      'Brudd på rutiner',
      'Manglende rutiner',
      'Menneskelig svikt',
      'Teknisk svikt',
      'Annet',
    ]);
  
    const relationsForPeopleInvolvedOptions: IChoiceGroupOption[] = options([
      'Ansatt/innleid',
      'NAV-bruker',
      'Annet',
    ]);

    return (<form onSubmit={this.sendForm}>
      <Stack tokens={{ childrenGap: 20}}>
        <TextField 
          label='Hva har skjedd? Beskriv hendelsen, hvorfor dette skjedde og hvilke konsekvenser hendelsen kan få eller har fått.'
          value={this.state.incidentDescription}
          onChange={(_, val) => this.setState({incidentDescription: val})}
          {...longTextFieldProps}
          required
          onGetErrorMessage={(value: string): string => value ? '' : 'Du må fylle ut dette feltet.'}
          validateOnLoad={false}
          validateOnFocusOut
        />
        <TextField
          label='Har du forslag til tiltak for å unngå at noe slik skjer igjen?'
          value={this.state.suggestedActions}
          onChange={(_, val) => this.setState({suggestedActions: val})}
          {...longTextFieldProps}
        />
        <Stack>
          <ChoiceGroup 
            label='Hva gjelder hendelsen?'
            options={categoryOptions}
            selectedKey={this.state.category}
            onChange={(_, opt) => this.setState({category: opt.key as string})}
            required
            />
          <TextField
            label='Du valgte «annet». Vennligst spesifiser:'
            value={this.state.categoryOther}
            onChange={(_, val) => this.setState({categoryOther: val})}
            disabled={this.state.category !== strings.IncidentCategoryOther}
            {...shortTextFieldProps}
            required={this.state.category === strings.IncidentCategoryOther}
          />
        </Stack>
        <Stack>
          <h2>Tilleggsopplysninger</h2>
          <p>For å kunne behandle din innsending på riktig måte, trenger vi noe mer informasjon. Er du usikker på hva du skal skrive, kan du la være å fylle inn de feltene.</p>
        </Stack>
        <DateTimePicker 
          label='Når skjedde/startet hendelsen, eller når ble den oppdaget?'
          value={this.state.incidentDate}
          onChange={val => this.setState({incidentDate: val})}
          {...dateTimeLocalizationProps}
        />
        <TextField 
          label='Hvor skjedde hendelsen?'
          description='Enhet / Geografisk lokasjon'
          value={this.state.incidentLocation}
          onChange={(_, val) => this.setState({incidentLocation: val})}
          {...shortTextFieldProps}
        />
        {this.state.category === strings.IncidentCategoryPrivacy && <>
          <Stack>
            <h3>Tilleggsspørsmål for hendelser knyttet til personvern</h3>
            <Checkbox label='Har personopplysninger havnet på avveie?' checked={this.state.personalInfoLost} onChange={(ev, checked) => this.setState({personalInfoLost: checked})}/>
          </Stack>
          { this.state.personalInfoLost && <>
            <p>Når personopplysninger er på avveie, trenger vi ekstra informasjon som skal rapporteres til Datatilsynet. Fyll inn så godt du kan.</p>
            <DatePicker 
              label='Hvor lenge varte hendelsen? (til hvilken dato)'
              value={this.state.incidentToDate}
              onSelectDate={val => this.setState({incidentToDate: val})}
              {...dateLocalizationProps}
            />
            <TextField
              label='Hvem er de berørte?'
              description='Oppgi navn og personnummer. Ett per linje.'
              value={this.state.peopleInvolved}
              onChange={(_, val) => this.setState({peopleInvolved: val})}
              {...longTextFieldProps}
            />
            <ChoiceGroup 
              label='Hovedårsak'
              options={incidentMainCauseOptions}
              selectedKey={this.state.incidentMainCause}
              onChange={(_, val) => this.setState({incidentMainCause: val.key})}
              {...choiceGroupProps}
            />
            <ChoiceGroup 
              label='Hvilken relasjon har virksomheten til de personene som er berørt av hendelsen?'
              options={relationsForPeopleInvolvedOptions}
              selectedKey={this.state.relationsForPeopleInvolved}
              onChange={(_, val) => this.setState({relationsForPeopleInvolved: val.key})}
              {...choiceGroupProps}
            />
            { this.state.relationsForPeopleInvolved &&
              this.state.relationsForPeopleInvolved === relationsForPeopleInvolvedOptions[2].key &&
              <TextField
                label='Du valgte «annet». Vennligst spesifiser:'
                value={this.state.relationsForPeopleInvolvedOther}
                onChange={(_, val) => this.setState({relationsForPeopleInvolvedOther: val})}
                {...shortTextFieldProps}
              />
            }
          </>}
        </>}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton 
            text='Send inn skjema'
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
        <Toggle
          label='Feilsøk'
          onChange={(ev, checked) => this.setState({debug: checked})}
        />
        { this.state.debug && <>
        <Stack tokens={{ childrenGap: 10 }}>
          <TextField
            label='Skjemadata'
            value={JSON.stringify({...this._getFormFields(), reporterID: this.props.context.pageContext.user.loginName}, undefined, 2)}
            readOnly multiline autoAdjustHeight
          />
          <DefaultButton
            text='Slett skjemadata fra nettleseren'
            onClick={this._resetState}
          />
        </Stack>
        </>}
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

  private _resetState = () => {
    PnpStorage.session.delete(PnpStorageKey);
    this.setState(DefatultState);
  }

  protected sendForm = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!this.props.azureFunctionUrl || !this.props.azureFunctionCode) {
      this.setState({
        hasError: true,
        errorMessage: 'Mangler url eller token til Azure-API. Dette må legges inn i nettdelens innstillinger.',
      });
      return;
    }
    this.setState({sending: true});
    const body = JSON.stringify(this._getFormFields());
    const headers: Headers = new Headers({
      'Content-Type': 'application/json',
      'X-Functions-Key': this.props.azureFunctionCode.trim(),
    });
    const httpClientOptions: IHttpClientOptions = {body, headers, };
    try {
      const response: HttpClientResponse = await this.props.context.httpClient.post(
        this.props.azureFunctionUrl.trim(),
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
      'debug',
    ].forEach(k => delete fields[k]);
    return fields;
  }

}
