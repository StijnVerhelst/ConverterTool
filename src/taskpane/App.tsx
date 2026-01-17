/**
 * Main App Component
 *
 * Tab-based layout for the ConverterTool Add-in.
 * Tabs: Data | Template | Preview | Export
 */

import { useState, useCallback } from "react";
import {
  TabList,
  Tab,
  SelectTabEvent,
  SelectTabData,
  makeStyles,
  tokens,
  Title3,
} from "@fluentui/react-components";
import {
  TableSimple24Regular,
  DocumentText24Regular,
  DocumentEdit24Regular,
  Eye24Regular,
  ArrowDownload24Regular,
} from "@fluentui/react-icons";

import { DataTab } from "./components/DataTab";
import { OutputsTab } from "./components/OutputsTab";
import { TemplateTab } from "./components/TemplateTab";
import { PreviewTab } from "./components/PreviewTab";
import { ExportTab } from "./components/ExportTab";

import type { AddInTemplate } from "@/types";
import type { TableInfo } from "@/services";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  header: {
    padding: "12px 16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  title: {
    margin: 0,
    color: tokens.colorNeutralForeground1,
  },
  tabList: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  content: {
    flex: 1,
    overflow: "auto",
    padding: "16px",
  },
});

type TabValue = "data" | "outputs" | "template" | "preview" | "export";

export function App() {
  const styles = useStyles();
  const [selectedTab, setSelectedTab] = useState<TabValue>("data");
  const [template, setTemplate] = useState<AddInTemplate | null>(null);
  const [tables, setTables] = useState<TableInfo[]>([]);

  const handleTabSelect = useCallback(
    (_event: SelectTabEvent, data: SelectTabData) => {
      setSelectedTab(data.value as TabValue);
    },
    []
  );

  const handleTemplateChange = useCallback((newTemplate: AddInTemplate | null) => {
    setTemplate(newTemplate);
  }, []);

  const handleTablesChange = useCallback((newTables: TableInfo[]) => {
    setTables(newTables);
  }, []);

  const renderContent = () => {
    switch (selectedTab) {
      case "data":
        return (
          <DataTab
            tables={tables}
            onTablesChange={handleTablesChange}
            template={template}
            onTemplateChange={handleTemplateChange}
          />
        );
      case "outputs":
        return (
          <OutputsTab
            template={template}
            onTemplateChange={handleTemplateChange}
            tables={tables}
          />
        );
      case "template":
        return (
          <TemplateTab
            template={template}
            tables={tables}
            onTemplateChange={handleTemplateChange}
          />
        );
      case "preview":
        return <PreviewTab template={template} />;
      case "export":
        return <ExportTab template={template} />;
      default:
        return null;
    }
  };

  return (
    <div className={styles.container}>
      <header className={styles.header}>
        <Title3 className={styles.title}>ConverterTool</Title3>
      </header>

      <TabList
        className={styles.tabList}
        selectedValue={selectedTab}
        onTabSelect={handleTabSelect}
        size="small"
      >
        <Tab value="data" icon={<TableSimple24Regular />}>
          Data
        </Tab>
        <Tab value="outputs" icon={<DocumentEdit24Regular />}>
          Outputs
        </Tab>
        <Tab value="template" icon={<DocumentText24Regular />}>
          Template
        </Tab>
        <Tab value="preview" icon={<Eye24Regular />}>
          Preview
        </Tab>
        <Tab value="export" icon={<ArrowDownload24Regular />}>
          Export
        </Tab>
      </TabList>

      <main className={styles.content}>{renderContent()}</main>
    </div>
  );
}
