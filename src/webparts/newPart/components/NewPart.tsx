import * as React from 'react';
import styles from './NewPart.module.scss';
import { ITestProps } from './INewPartProps';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";

//import { PrimaryButton } from '@fluentui/react/lib/Button';
type PropsTypes = {
  Links: {
    Url: string
  }
  Title: string
}

type MyState = {
  title: string;
  link: string;
}


export default class NewPart extends React.Component<ITestProps, MyState> {
  constructor(props: ITestProps) {
    super(props);
    this.state = {
      title: "",
      link: ""
    }
  }

  private getListItems = async () => {
      const sp = spfi(this.props.context.pageContext.site.absoluteUrl).using(SPFx(this.props.context))
      const list = sp.web.getList(this.props.context.pageContext.web.serverRelativeUrl + "/lists/" + 'Useful links');
      const elem = list.items();
      elem.then((response) => {
        this._renderList(response);
      console.log(response)

       })
       .catch(() => {});
  }

  private createListItem = async () => {
    const { title, link } = this.state;
    try {
      const sp = spfi(this.props.context.pageContext.site.absoluteUrl).using(SPFx(this.props.context))
      const list = sp.web.getList(this.props.context.pageContext.web.serverRelativeUrl + "/lists/" + 'Useful links');
      const newItem = await list.items.add({
        Title: title,
        Links: {
          Description: link,
          Url: link
        }

      });
      console.log(newItem)
    } catch (error) {

      console.log('Error adding new item:');

    }
  }

  private onSubmit = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (this.state.title.length < 2 || this.state.link.length < 2) return;
    this.createListItem();
    this.getListItems();
  }

  public onDelete = () => {
    console.log("delete")
}

  private _renderList(items: PropsTypes[]): void {
    let html: string = '';
    items.forEach((item: PropsTypes) => {
      html += `
        <li class="${styles.list_group}">
            <a href="${item.Links.Url}">${item.Title}</a>
            <input type="text" class="${styles.list_group_item}" name="input" value="${item.Title}" maxlength="15"  disabled="true">
            <div class='${styles.list_group_button}'>
                <button type="button"
                          class="${styles.btn_edit}"
                          onClick={onToggleProp}>Edit</button>
                  <button type="button"
                          class="${styles.btn_delete}"
                          onClick=${this.onDelete}>Delete
                  </button> 
            </div>
       </li>`;
       
    });
     
    const listContainer: Element = document.querySelector('#spListContainer') as HTMLElement;
    listContainer.innerHTML = html;
  }
  
  public render(): React.ReactElement<ITestProps> {
    const {
      //description,
      isDarkTheme,
      hasTeamsContext,
      //listIndex
    } = this.props;


    return (
      <section className={`${styles.newPart} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.table}>
          <ul className={styles.img}>
            <li className={styles.img_item}>
              <img className={styles.img_item_img} alt="" src={isDarkTheme ? require('../assets/career.png') : require('../assets/career.png')} />
            </li>
            <li className={styles.img_item}>
              <img className={styles.img_item_img} alt="" src={isDarkTheme ? require('../assets/success.png') : require('../assets/success.png')} />
            </li>
            <li className={styles.img_item}>
              <img className={styles.img_item_img} alt="" src={isDarkTheme ? require('../assets/knowelledge.png') : require('../assets/knowelledge.png')} />
            </li>
            <li className={styles.img_item}>
              <img className={styles.img_item_img} alt="" src={isDarkTheme ? require('../assets/vacancies.png') : require('../assets/vacancies.png')} />
            </li>
          </ul>
        </div>
        <div className={styles.links} >
          <ul id="spListContainer" className={styles.list}>
          </ul>
        </div>
        <div className={styles.add_form}>
          <h3>Add element</h3>
          <form className={styles.forms} onSubmit={this.onSubmit} >
            <input type="text"
              className={styles.form_control}
              placeholder="Item name"
              name="title"
              value={this.state.title}
              onChange={(e: React.ChangeEvent<HTMLInputElement>): void => this.setState({ title: e.target.value })} />
            <input type="text"
              className={styles.form_control}
              placeholder="Link"
              name="link"
              value={this.state.link}
              onChange={(e: React.ChangeEvent<HTMLInputElement>): void => this.setState({ link: e.target.value })} />

            <button type="submit" className={styles.btn_edit}>Add</button>
          </form>
        </div>
      </section>
    );

  }
}
