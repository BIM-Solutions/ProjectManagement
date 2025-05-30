import * as React from "react";
import { makeStyles, tokens, Tooltip } from "@fluentui/react-components";
import {
  PersonCircle32Regular,
  Board20Filled,
  Board20Regular,
  PersonClock20Filled,
  PersonClock20Regular,
  bundleIcon,
} from "@fluentui/react-icons";
import {
  AppItem,
  Hamburger,
  NavDrawer,
  NavDrawerBody,
  NavDrawerHeader,
  NavItem,
} from "@fluentui/react-nav-preview";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
interface INavigationProps {
  context: WebPartContext;
  currentPage: string;
  // Define any props you need for the Navigation component here
}

const useStyles = makeStyles({
  root: {
    overflow: "hidden",
    display: "flex",
    flex: 1,
    minHeight: 0,
  },
  nav: {
    minWidth: "160px",
    width: "160px",
  },
  content: {
    flex: "1",
    padding: "16px",
    display: "grid",
    justifyContent: "flex-start",
    alignItems: "flex-start",
  },
  field: {
    display: "flex",
    marginTop: "4px",
    marginLeft: "8px",
    flexDirection: "column",
    gridRowGap: tokens.spacingVerticalS,
  },
});

const Navigation: React.FC<INavigationProps> = ({ context, currentPage }) => {
  const [isOpen, setIsOpen] = React.useState(true);
  const Dashboard = bundleIcon(Board20Filled, Board20Regular);
  const Person = bundleIcon(PersonClock20Filled, PersonClock20Regular);
  const styles = useStyles();
  const [siteRoot, setSiteRoot] = React.useState<string>("");

  React.useEffect(() => {
    const fetchSiteRoot = async (): Promise<void> => {
      // Replace 'sp' with your actual SPFI instance if needed
      const sp = spfi().using(SPFx(context));
      const webData = await sp.web.select("ServerRelativeUrl")();
      if (webData && webData.ServerRelativeUrl) {
        setSiteRoot(webData.ServerRelativeUrl.replace(/\/$/, ""));
      }
    };
    fetchSiteRoot().catch(console.error);
  }, []);

  const linkHome = siteRoot;
  const linkResource = siteRoot + "/SitePages/resourcing.aspx";
  const linkCofw = siteRoot + "/SitePages/cofw.aspx";

  return (
    <div>
      {isOpen && (
        <NavDrawer
          defaultSelectedValue={currentPage}
          defaultSelectedCategoryValue=""
          open={isOpen}
          type="inline"
          multiple={true}
          className={styles.nav}
          style={{ height: "100%", minWidth: "200px" }}
        >
          <NavDrawerHeader>
            <Tooltip content="Close Navigation" relationship="label">
              <Hamburger onClick={() => setIsOpen(!isOpen)} />
            </Tooltip>
          </NavDrawerHeader>

          <NavDrawerBody>
            <AppItem icon={<PersonCircle32Regular />} as="a" href={siteRoot}>
              BIM Team
            </AppItem>
            <NavItem href={linkHome} icon={<Dashboard />} value="1">
              Projects
            </NavItem>
            <NavItem href={linkResource} icon={<Person />} value="2">
              Resourcing
            </NavItem>
            <NavItem href={linkCofw} icon={<Dashboard />} value="3">
              CoFW
            </NavItem>
          </NavDrawerBody>
        </NavDrawer>
      )}
      {!isOpen && (
        <Tooltip content="Open Navigation" relationship="label">
          <Hamburger
            style={{ height: "32px" }}
            onClick={() => setIsOpen(!isOpen)}
          />
        </Tooltip>
      )}
    </div>
  );
};

export default Navigation;
