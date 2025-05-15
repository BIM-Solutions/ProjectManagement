
import * as React from 'react';
import {
    makeStyles,
    tokens,
    Tooltip,
  } from '@fluentui/react-components';
  import {
    PersonCircle32Regular,
    Board20Filled,
    Board20Regular,
    PersonClock20Filled,
    PersonClock20Regular,
    bundleIcon,
  } from '@fluentui/react-icons';
  import {
    AppItem,
    Hamburger,
    NavDrawer,
    NavDrawerBody,
    NavDrawerHeader,
    NavItem,
  } from '@fluentui/react-nav-preview';
interface INavigationProps {

    // Define any props you need for the Navigation component here
}

const useStyles = makeStyles({
    root: {
      overflow: 'hidden',
      display: 'flex',
      flex: 1,
      minHeight: 0, 
    },
    nav: {
      minWidth: '160px',
      width: '160px',
    },
    content: {
      flex: '1',
      padding: '16px',
      display: 'grid',
      justifyContent: 'flex-start',
      alignItems: 'flex-start',
    },
    field: {
      display: 'flex',
      marginTop: '4px',
      marginLeft: '8px',
      flexDirection: 'column',
      gridRowGap: tokens.spacingVerticalS,
    },
  });

const Navigation: React.FC<INavigationProps> = (props) => {
    const [isOpen, setIsOpen] = React.useState(true);
    const Dashboard = bundleIcon(Board20Filled, Board20Regular);
    const Person = bundleIcon(PersonClock20Filled, PersonClock20Regular);
    const styles = useStyles();
    const linkDestination = 'https://contoso.sharepoint.com/sites/ContosoHR/Pages/Home.aspx';
    

    return( 
        <div>
            {isOpen && (
                <NavDrawer
                defaultSelectedValue="1"
                defaultSelectedCategoryValue=""
                open={isOpen}
                type="inline"
                multiple={true}
                className={styles.nav}
                style={{ height: '100%', minWidth: '200px' }}
                >
                <NavDrawerHeader>
                <Tooltip content="Close Navigation" relationship="label">
                    <Hamburger onClick={() => setIsOpen(!isOpen)} />
                </Tooltip>
                </NavDrawerHeader>
        
                <NavDrawerBody>
                <AppItem icon={<PersonCircle32Regular />} as="a" href={linkDestination}>
                    BIM Team
                </AppItem>
                <NavItem href={linkDestination} icon={<Dashboard />} value="1">
                    Projects
                </NavItem>
                <NavItem href={linkDestination} icon={<Person />} value="2">
                    Resourcing
                </NavItem>
                <NavItem href={linkDestination} icon={<Dashboard />} value="3">
                    CoFW
                </NavItem>
                </NavDrawerBody>
            </NavDrawer>
            )}
            {!isOpen && (
                <Tooltip content="Open Navigation" relationship="label">
                <Hamburger style={{ height: '32px' }} onClick={() => setIsOpen(!isOpen)} />
                </Tooltip>
            )}
        </div>
    
    );  
};

export default Navigation;