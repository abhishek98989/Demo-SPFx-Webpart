import React, { useState, useEffect } from 'react';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import './style.css';
interface Banner {
  Title: string;
  CategoryDescription: string;
  Alert_x0020_Type: 'Info-Light-Green' | 'Success-Green' | 'Warning-Yellow' | 'Alert-Red';
  Alert_x0020_Status: 'Active' | 'Inactive';
}

interface SharePointBannersProps {
  context: any; // SPFx context
  listName?: string;
}

const SharePointBanners= (props:any) => {
  const [banners, setBanners] = useState<Banner[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const fetchBanners = async () => {
      try {
        setLoading(true);
        setError(null);

        // Initialize PnP SP with SPFx context
        const sp = spfi('https://vaughnconstruction.sharepoint.com').using(SPFx(props?.context));

        // Fetch active banners from SharePoint list
        const items = await sp.web.lists
          .getById('274cad54-c0ab-451e-91f2-4ecae3742a9b')
          .items
          .select('Title', 'CategoryDescription', 'Alert_x0020_Type', 'Alert_x0020_Status')
          .filter("Alert_x0020_Status eq 'Active'")
          .orderBy('Modified', false)
          .top(100)();

        setBanners(items as Banner[]);
      } catch (err) {
        console.error('Error fetching banners:', err);
        setError('Failed to load banners. Please try again later.');
      } finally {
        setLoading(false);
      }
    };

    if (props?.context) {
      fetchBanners();
    }
  }, []);

  
const getAlertClassName = (alertType: Banner['Alert_x0020_Type']) => {
    switch (alertType) {
      case 'Info-Light-Green':
        return 'sp-banner-info-light-green';
      case 'Success-Green':
        return 'sp-banner-success-green';
      case 'Warning-Yellow':
        return 'sp-banner-warning-yellow';
      case 'Alert-Red':
        return 'sp-banner-alert-red';
      default:
        return 'sp-banner-default';
    }
  };

  return (
    <>
    {banners?.length>0&&
     <div className="sp-banners-container">
      {banners?.map((banner, index) => {
        const alertClass = getAlertClassName(banner?.Alert_x0020_Type);

        return (
          <div key={index} className={`sp-banner ${alertClass}`}>
            <div className="sp-banner-content">
              <div className="sp-icon-wrapper">
              </div>
              <div className="sp-text-content">
                <h3 className="sp-banner-title">{banner?.Title}</h3>
                <p className="sp-banner-description" dangerouslySetInnerHTML={{ __html:banner?.CategoryDescription }} ></p>
              </div>
            </div>
          </div>
        );
      })}
    </div>
    }
    </>
   
  );
};

export default SharePointBanners;