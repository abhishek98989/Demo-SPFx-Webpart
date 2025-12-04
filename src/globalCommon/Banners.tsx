import React, { useState, useEffect } from 'react';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import './style.css';

interface Banner {
  Id: number;
  Title: string;
  CategoryDescription: string;
  Alert_x0020_Type: 'Info-Light-Green' | 'Success-Green' | 'Warning-Yellow' | 'Alert-Red';
  Alert_x0020_Status: 'Active' | 'Inactive';
  Modified: string; // ISO date string from SharePoint
}

const STORAGE_KEY_PREFIX = 'spBannerDismiss_';

const SharePointBanners = (props: any) => {
  const [banners, setBanners] = useState<Banner[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // --- LocalStorage helpers ---

  const getStorageKey = (bannerId: number) =>
    `${STORAGE_KEY_PREFIX}${bannerId}`;

  const shouldShowBanner = (banner: Banner): boolean => {
    if (typeof window === 'undefined') return true; // safety for non-browser

    try {
      const key = getStorageKey(banner.Id);
      const raw = window.localStorage.getItem(key);
      if (!raw) {
        // never dismissed
        return true;
      }

      const data = JSON.parse(raw) as {
        lastDismissed: string;
        lastModified: string;
      };

      // If the banner was modified AFTER it was dismissed, show it again
      if (!data.lastModified || data.lastModified !== banner.Modified) {
        return true;
      }

      // Same Modified as when dismissed
      if (banner.Alert_x0020_Type === 'Alert-Red') {
        // For red banners, they reappear after 12 hours if still active
        const lastDismissedDate = new Date(data.lastDismissed);
        const now = new Date();

        if (isNaN(lastDismissedDate.getTime())) {
          // invalid date in storage: show it
          return true;
        }

        const diffMs = now.getTime() - lastDismissedDate.getTime();
        const diffHours = diffMs / (1000 * 60 * 60);

        // Reappear after 12 hours
        return diffHours >= 12;
      }

      // For other banners: stay hidden as long as Modified is unchanged
      return false;
    } catch (e) {
      // On any error, fail open (show the banner)
      return true;
    }
  };

  const handleDismiss = (banner: Banner) => {
    // Update UI
    setBanners((prev) => prev.filter((b) => b.Id !== banner.Id));

    // Persist dismissal
    if (typeof window !== 'undefined') {
      try {
        const key = getStorageKey(banner.Id);
        const payload = {
          lastDismissed: new Date().toISOString(),
          lastModified: banner.Modified,
        };
        window.localStorage.setItem(key, JSON.stringify(payload));
      } catch (e) {
        // ignore storage errors
        console.error('Error writing banner dismiss info to localStorage', e);
      }
    }
  };

  // --- Fetch banners from SharePoint ---

  useEffect(() => {
    const fetchBanners = async () => {
      try {
        setLoading(true);
        setError(null);

        const sp = spfi('https://vaughnconstruction.sharepoint.com').using(
          SPFx(props?.context)
        );

        const items: any[] = await sp.web.lists
          .getById('274cad54-c0ab-451e-91f2-4ecae3742a9b')
          .items.select(
            'Id',
            'Title',
            'CategoryDescription',
            'Alert_x0020_Type',
            'Alert_x0020_Status',
            'Modified'
          )
          .filter("Alert_x0020_Status eq 'Active'")
          .orderBy('Modified', false)
          .top(100)();

        // Filter based on localStorage rules BEFORE putting into state
        const mapped = items as Banner[];
        const visible = mapped.filter((banner) => shouldShowBanner(banner));

        setBanners(visible);
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
  }, [props?.context]); // re-run if context changes

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

  if (loading) {
    return null; // or a loader if you want
  }

  if (error) {
    return null; // or show error if needed
  }

  return (
    <>
      {banners && banners.length > 0 && (
        <div className="sp-banners-container">
          {banners.map((banner) => {
            const alertClass = getAlertClassName(banner.Alert_x0020_Type);

            return (
              <div key={banner.Id} className={`sp-banner ${alertClass}`}>
                {/* Close / cross icon */}
                <button
                  className="sp-banner-close"
                  aria-label="Dismiss notification"
                  onClick={() => handleDismiss(banner)}
                >
                  ×
                </button>

                <div className="sp-banner-content">
                  <div className="sp-icon-wrapper"></div>
                  <div className="sp-text-content">
                    <h3 className="sp-banner-title">{banner.Title}</h3>
                    <p
                      className="sp-banner-description"
                      dangerouslySetInnerHTML={{
                        __html: banner.CategoryDescription,
                      }}
                    ></p>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </>
  );
};

export default SharePointBanners;
