import { SocialAccount } from '../contexts/SocialAccountsContext';

export interface PostData {
  content: string;
  mediaUrl?: string;
  mediaType?: 'image' | 'video';
  scheduledTime?: Date;
}

export interface PublishResult {
  success: boolean;
  platform: string;
  postId?: string;
  url?: string;
  error?: string;
}

class SocialMediaService {
  // Twitter/X API integration
  async publishToTwitter(account: SocialAccount, postData: PostData): Promise<PublishResult> {
    try {
      // In a real implementation, this would use the Twitter API v2
      const response = await fetch('https://api.twitter.com/2/tweets', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${account.accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          text: postData.content,
          ...(postData.mediaUrl && { media: { media_ids: [await this.uploadMedia(account, postData.mediaUrl)] } })
        })
      });

      if (!response.ok) {
        throw new Error(`Twitter API error: ${response.status}`);
      }

      const data = await response.json();
      
      return {
        success: true,
        platform: 'twitter',
        postId: data.data.id,
        url: `https://twitter.com/${account.username}/status/${data.data.id}`
      };
    } catch (error: any) {
      return {
        success: false,
        platform: 'twitter',
        error: error.message || 'Failed to publish to Twitter'
      };
    }
  }

  // Facebook API integration
  async publishToFacebook(account: SocialAccount, postData: PostData): Promise<PublishResult> {
    try {
      // Facebook Graph API
      const pageId = await this.getFacebookPageId(account);
      const endpoint = `https://graph.facebook.com/v18.0/${pageId}/feed`;
      
      const body: any = {
        message: postData.content,
        access_token: account.accessToken
      };

      if (postData.mediaUrl) {
        if (postData.mediaType === 'image') {
          body.link = postData.mediaUrl;
        } else if (postData.mediaType === 'video') {
          // For videos, use the videos endpoint
          return await this.publishVideoToFacebook(account, postData);
        }
      }

      const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(body)
      });

      if (!response.ok) {
        throw new Error(`Facebook API error: ${response.status}`);
      }

      const data = await response.json();
      
      return {
        success: true,
        platform: 'facebook',
        postId: data.id,
        url: `https://facebook.com/${data.id}`
      };
    } catch (error: any) {
      return {
        success: false,
        platform: 'facebook',
        error: error.message || 'Failed to publish to Facebook'
      };
    }
  }

  // Instagram API integration
  async publishToInstagram(account: SocialAccount, postData: PostData): Promise<PublishResult> {
    try {
      // Instagram Basic Display API / Instagram Graph API
      const instagramAccountId = await this.getInstagramAccountId(account);
      
      if (postData.mediaType === 'image') {
        return await this.publishImageToInstagram(account, postData, instagramAccountId);
      } else if (postData.mediaType === 'video') {
        return await this.publishVideoToInstagram(account, postData, instagramAccountId);
      } else {
        throw new Error('Instagram requires media content');
      }
    } catch (error: any) {
      return {
        success: false,
        platform: 'instagram',
        error: error.message || 'Failed to publish to Instagram'
      };
    }
  }

  // LinkedIn API integration
  async publishToLinkedIn(account: SocialAccount, postData: PostData): Promise<PublishResult> {
    try {
      // LinkedIn API v2
      const personUrn = await this.getLinkedInPersonUrn(account);
      
      const shareData: any = {
        author: personUrn,
        lifecycleState: 'PUBLISHED',
        specificContent: {
          'com.linkedin.ugc.ShareContent': {
            shareCommentary: {
              text: postData.content
            },
            shareMediaCategory: postData.mediaUrl ? 'IMAGE' : 'NONE'
          }
        },
        visibility: {
          'com.linkedin.ugc.MemberNetworkVisibility': 'PUBLIC'
        }
      };

      if (postData.mediaUrl) {
        const mediaUrn = await this.uploadMediaToLinkedIn(account, postData.mediaUrl);
        shareData.specificContent['com.linkedin.ugc.ShareContent'].media = [{
          status: 'READY',
          description: {
            text: 'Shared media'
          },
          media: mediaUrn,
          title: {
            text: 'Media'
          }
        }];
      }

      const response = await fetch('https://api.linkedin.com/v2/ugcPosts', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${account.accessToken}`,
          'Content-Type': 'application/json',
          'X-Restli-Protocol-Version': '2.0.0'
        },
        body: JSON.stringify(shareData)
      });

      if (!response.ok) {
        throw new Error(`LinkedIn API error: ${response.status}`);
      }

      const data = await response.json();
      
      return {
        success: true,
        platform: 'linkedin',
        postId: data.id,
        url: `https://linkedin.com/feed/update/${data.id}`
      };
    } catch (error: any) {
      return {
        success: false,
        platform: 'linkedin',
        error: error.message || 'Failed to publish to LinkedIn'
      };
    }
  }

  // Generic publish method that routes to platform-specific methods
  async publishPost(account: SocialAccount, postData: PostData): Promise<PublishResult> {
    // Validate account status
    if (!account.isConnected || account.status !== 'active') {
      return {
        success: false,
        platform: account.platform,
        error: 'Account is not connected or active'
      };
    }

    // Check if token is expired
    if (account.expiresAt && account.expiresAt < new Date()) {
      return {
        success: false,
        platform: account.platform,
        error: 'Access token has expired. Please reconnect your account.'
      };
    }

    switch (account.platform) {
      case 'twitter':
        return await this.publishToTwitter(account, postData);
      case 'facebook':
        return await this.publishToFacebook(account, postData);
      case 'instagram':
        return await this.publishToInstagram(account, postData);
      case 'linkedin':
        return await this.publishToLinkedIn(account, postData);
      default:
        return {
          success: false,
          platform: account.platform,
          error: `Publishing to ${account.platform} is not yet supported`
        };
    }
  }

  // Batch publish to multiple platforms
  async publishToMultiplePlatforms(
    accounts: SocialAccount[], 
    postData: PostData
  ): Promise<PublishResult[]> {
    const publishPromises = accounts.map(account => this.publishPost(account, postData));
    return await Promise.all(publishPromises);
  }

  // Helper methods (these would contain actual API implementation details)
  private async uploadMedia(account: SocialAccount, mediaUrl: string): Promise<string> {
    // Implementation for uploading media to the respective platform
    return 'mock_media_id';
  }

  private async getFacebookPageId(account: SocialAccount): Promise<string> {
    // Get Facebook page ID for posting
    return 'mock_page_id';
  }

  private async publishVideoToFacebook(account: SocialAccount, postData: PostData): Promise<PublishResult> {
    // Facebook video publishing implementation
    return {
      success: true,
      platform: 'facebook',
      postId: 'mock_video_post_id'
    };
  }

  private async getInstagramAccountId(account: SocialAccount): Promise<string> {
    // Get Instagram business account ID
    return 'mock_instagram_id';
  }

  private async publishImageToInstagram(
    account: SocialAccount, 
    postData: PostData, 
    accountId: string
  ): Promise<PublishResult> {
    // Instagram image publishing implementation
    return {
      success: true,
      platform: 'instagram',
      postId: 'mock_instagram_post_id'
    };
  }

  private async publishVideoToInstagram(
    account: SocialAccount, 
    postData: PostData, 
    accountId: string
  ): Promise<PublishResult> {
    // Instagram video publishing implementation
    return {
      success: true,
      platform: 'instagram',
      postId: 'mock_instagram_video_id'
    };
  }

  private async getLinkedInPersonUrn(account: SocialAccount): Promise<string> {
    // Get LinkedIn person URN
    return 'urn:li:person:mock_person_id';
  }

  private async uploadMediaToLinkedIn(account: SocialAccount, mediaUrl: string): Promise<string> {
    // Upload media to LinkedIn and return URN
    return 'urn:li:digitalmediaAsset:mock_media_id';
  }

  // Schedule post for later publishing
  async schedulePost(
    accounts: SocialAccount[], 
    postData: PostData, 
    scheduledTime: Date
  ): Promise<{ success: boolean; scheduleId?: string; error?: string }> {
    try {
      // In a real implementation, this would store the scheduled post in a database
      // and use a job queue (like Bull, Agenda, or cloud functions) to publish at the scheduled time
      
      const scheduleId = `schedule_${Date.now()}`;
      
      // Store in localStorage for demo purposes
      const scheduledPosts = JSON.parse(localStorage.getItem('scheduled_posts') || '[]');
      scheduledPosts.push({
        id: scheduleId,
        accounts: accounts.map(a => a.id),
        postData,
        scheduledTime: scheduledTime.toISOString(),
        status: 'scheduled'
      });
      localStorage.setItem('scheduled_posts', JSON.stringify(scheduledPosts));

      return {
        success: true,
        scheduleId
      };
    } catch (error: any) {
      return {
        success: false,
        error: error.message || 'Failed to schedule post'
      };
    }
  }

  // Get scheduled posts
  getScheduledPosts(): any[] {
    try {
      return JSON.parse(localStorage.getItem('scheduled_posts') || '[]');
    } catch {
      return [];
    }
  }

  // Cancel scheduled post
  async cancelScheduledPost(scheduleId: string): Promise<boolean> {
    try {
      const scheduledPosts = this.getScheduledPosts();
      const updatedPosts = scheduledPosts.filter(post => post.id !== scheduleId);
      localStorage.setItem('scheduled_posts', JSON.stringify(updatedPosts));
      return true;
    } catch {
      return false;
    }
  }
}

export const socialMediaService = new SocialMediaService();
export default socialMediaService;