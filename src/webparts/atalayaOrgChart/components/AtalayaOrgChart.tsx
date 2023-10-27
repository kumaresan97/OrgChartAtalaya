import * as React from "react";
// import styles from './AtalayaOrgChart.module.scss';
import { IAtalayaOrgChartProps } from "./IAtalayaOrgChartProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { graph } from "@pnp/graph/presets/all";
import OrgChart from "./OrgChart";
import "./style.css";

export default class AtalayaOrgChart extends React.Component<
  IAtalayaOrgChartProps,
  {}
> {
  constructor(props: IAtalayaOrgChartProps) {
    super(props);
    graph.setup({
      spfxContext: this.props.context,
    });
  }

  public render(): React.ReactElement<IAtalayaOrgChartProps> {
    return <OrgChart context={this.props.context} />;
  }
}
