import * as React from 'react';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import styles from './GraphProfile.module.scss';
import { IGraphProfileProps } from './IGraphProfileProps';
import { IUserProfile, ISiteInfo, ISiteList } from '../Models/IGraphData';
import { graphErrorText } from '../Services/GraphService';

const columns: IColumn[] = [
  { key: 'name', name: 'List or library', fieldName: 'displayName', minWidth: 160, maxWidth: 260, isResizable: true },
  { key: 'template', name: 'Template', fieldName: 'template', minWidth: 110, maxWidth: 160 },
  {
    key: 'modified', name: 'Last modified', fieldName: 'lastModified', minWidth: 120,
    onRender: (item: ISiteList) => new Date(item.lastModified).toLocaleDateString('en-US')
  }
];

const GraphProfile: React.FC<IGraphProfileProps> = (props) => {
  const { service, profileSource, listCount } = props;
  const [me, setMe] = React.useState<IUserProfile | undefined>();
  const [photoUrl, setPhotoUrl] = React.useState<string | undefined>();
  const [site, setSite] = React.useState<ISiteInfo | undefined>();
  const [lists, setLists] = React.useState<ISiteList[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [profileError, setProfileError] = React.useState<string | undefined>();
  const [listError, setListError] = React.useState<string | undefined>();

  React.useEffect(() => {
    let cancelled = false;
    setLoading(true);

    // Profile + photo need only User.Read. Site + lists need Sites.Read.All.
    // Separate try/catch blocks, so a missing Sites.Read.All grant doesn't hide the profile.
    const loadProfile = async (): Promise<void> => {
      try {
        const [profile, photo] = await Promise.all([service.getMe(profileSource), service.getMyPhotoUrl()]);
        if (!cancelled) { setMe(profile); setPhotoUrl(photo); setProfileError(undefined); }
      } catch (error) {
        if (!cancelled) { setProfileError(graphErrorText(error)); }
      }
    };
    const loadLists = async (): Promise<void> => {
      try {
        const [siteInfo, siteLists] = await Promise.all([service.getSite(), service.getLists(listCount)]);
        if (!cancelled) { setSite(siteInfo); setLists(siteLists); setListError(undefined); }
      } catch (error) {
        if (!cancelled) { setListError(graphErrorText(error)); }
      }
    };

    Promise.all([loadProfile(), loadLists()])
      .then(() => { if (!cancelled) { setLoading(false); } })
      .catch(() => { if (!cancelled) { setLoading(false); } });

    return () => { cancelled = true; };
  }, [service, profileSource, listCount]);

  // Release the photo's object URL when it changes or the web part unloads.
  React.useEffect(() => () => { if (photoUrl) { URL.revokeObjectURL(photoUrl); } }, [photoUrl]);

  if (loading) {
    return <Spinner size={SpinnerSize.large} label="Calling Microsoft Graph..." />;
  }

  return (
    <section className={styles.graphProfile}>
      <h2 className={styles.title}>My Microsoft Graph profile</h2>
      {profileError && <MessageBar messageBarType={MessageBarType.error}>Profile: {profileError}</MessageBar>}
      {me && (
        <Persona
          imageUrl={photoUrl}
          text={me.displayName}
          secondaryText={me.jobTitle || me.userPrincipalName}
          tertiaryText={me.mail && me.mail !== me.userPrincipalName ? me.mail : undefined}
          optionalText={me.officeLocation}
          size={PersonaSize.size72}
        />
      )}
      <p className={styles.source}>
        Profile loaded with {profileSource === 'pnp' ? 'PnPjs graphfi()' : 'MSGraphClientV3'}
        {photoUrl ? ', photo from /me/photo/$value' : ', no profile photo (404)'}
      </p>

      <h3 className={styles.subtitle}>{site ? `Lists on ${site.displayName}` : 'Site lists'}</h3>
      {listError && <MessageBar messageBarType={MessageBarType.error}>Lists: {listError}</MessageBar>}
      {!listError && (
        <DetailsList
          items={lists}
          columns={columns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          compact={true}
        />
      )}
    </section>
  );
};

export default GraphProfile;