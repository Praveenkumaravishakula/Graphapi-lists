export interface IGraphProfileProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
import { GraphService } from '../Services/GraphService';
import { ProfileSource } from '../Models/IGraphData';

export interface IGraphProfileProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  service: GraphService;
  profileSource: ProfileSource;
  listCount: number;
}