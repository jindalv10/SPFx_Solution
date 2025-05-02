import { IMapArea } from "../../models/IMapArea";

export interface IImageMapperLandingPageProps {
  description: string;
  imageUrl: string;
  imageHeight: string;
  imageWidth: string;
  imageHorizontalPosition: string;
  imageVerticalPosition: string;
  scale: number;
  items: IMapArea[];
}
