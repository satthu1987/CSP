import * as React from 'react';
import styles from './About.module.scss';
import { Icon } from '@fluentui/react';

export interface IAboutProps {
  context: any;
}

interface IAboutState {
  expandedIds: { [key: string]: boolean };
}

export default class About extends React.Component<IAboutProps, IAboutState> {
  constructor(props: IAboutProps) {
    super(props);
    this.state = {
      expandedIds: {}
    };
  }

  private toggleSection = (id: string): void => {
    this.setState(prev => ({
      expandedIds: { ...prev.expandedIds, [id]: !prev.expandedIds[id] }
    }));
  };

  private renderCollapsibleSection = (id: string, title: string, content: string): JSX.Element => {
    const expanded = !!this.state.expandedIds[id];
    return (
      <div className={styles.collapsibleSection} key={id}>
        <button
          className={styles.sectionHeader}
          onClick={() => this.toggleSection(id)}
          aria-expanded={expanded}
        >
          <Icon
            iconName={expanded ? 'ChevronDown' : 'ChevronRight'}
            className={styles.chevron}
          />
          <span className={styles.sectionTitle}>{title}</span>
        </button>
        {expanded && (
          <div className={styles.sectionBody}>
            <p>{content}</p>
          </div>
        )}
      </div>
    );
  };

  public render(): JSX.Element {
    return (
      <main className={styles.main}>
        <div className={styles.breadcrumb}>Home › <strong>Learn More</strong></div>
        <div className={styles.heroContent}>
          <h1 className={styles.welcomeScript}>Welcome to</h1>
          <h2 className={styles.userName}>CSP Platform</h2>
        </div>

        <div className={styles.heroVisual}>
          <img
            src={require('../../assets/hero-image.png')}
            alt="Customer Satisfaction Program"
            className={styles.heroImage}
          />
          <div className={styles.heroTitle}>
            <span>CUSTOMER</span>
            <span>SATISFACTION</span>
            <span>PROGRAM</span>
          </div>
        </div>

        <section className={styles.informationSection}>
          <h2>Learn More</h2>
          {this.renderCollapsibleSection(
            'howItWorks',
            'How It Works',
            'The Customer Satisfaction Program is designed to gather feedback from our customers and convert it into actionable improvements. Our team systematically collects feedback, analyzes trends, and develops action plans to address customer concerns and enhance overall satisfaction.'
          )}
          {this.renderCollapsibleSection(
            'getInvolved',
            'Get Involved',
            'We encourage all team members to participate in the Customer Satisfaction Program. You can contribute by actively listening to customer feedback, sharing insights, and helping implement improvements. Visit our Resources section to learn how to get involved and make a difference in our customer relationships.'
          )}
          {this.renderCollapsibleSection(
            'measureInProgram',
            'Measure in the Program',
            'Our program measures success through multiple key performance indicators including customer satisfaction scores, feedback response times, action plan completion rates, and customer retention metrics. Regular reporting and data analysis help us track progress and identify areas for continuous improvement.'
          )}
          {this.renderCollapsibleSection(
            'impactedStory',
            'The Impacted Story',
            'Real-world customer feedback has led to significant operational improvements and enhanced service delivery. Several customers have reported improved satisfaction levels after their feedback was addressed through our action planning process. These success stories demonstrate the tangible impact of listening to and acting on customer input.'
          )}
        </section>
      </main>
    );
  }
}