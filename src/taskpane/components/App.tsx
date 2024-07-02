/* eslint-disable no-undef */
import * as React from "react";
//import { useState } from "react";
import { HfInference } from "@huggingface/inference";
//import { textGeneration } from "@huggingface/inference";
import { DefaultButton } from "@fluentui/react";
import Progress from "./Progress";
import { ChatCompletionRequestMessage, Configuration, OpenAIApi } from "openai";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  generatedText: string;
  generatedTextChinese: string;
  startText: string;
  startTextSave: string; //用来保存数据这样回来按键跳转回来可以继续编辑,恰好运用了页面切换刷新defaultvalue
  finalMailText: string;
  isLoading: boolean;
  isGenerateBusinessMailActive: boolean;
  isSummarizeMailActive: boolean;
  isTranslationMailActive: boolean;
  summary: string;
  translation: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props) {
    super(props);

    let isGenerateBusinessMailActive;
    let isSummarizeMailActive;
    let isTranslationMailActive;

    //get the current URL
    const url = window.location.href;
    console.log("URL: " + url);
    //check if the URL contains the parameter "generate"
    if (url.indexOf("compose") > -1) {
      console.log("Action: generate business mail");
      isGenerateBusinessMailActive = true;
      isSummarizeMailActive = false;
      isTranslationMailActive = false;
    }
    //check if the URL contains the parameter "summarize"
    if (url.indexOf("summary") > -1) {
      console.log("Action: summarize mail");
      isGenerateBusinessMailActive = false;
      isSummarizeMailActive = true;
      isTranslationMailActive = false;
    }
    //check if the URL contains the parameter "summarize"
    if (url.indexOf("translation") > -1) {
      console.log("Action: translate mail");
      isGenerateBusinessMailActive = false;
      isSummarizeMailActive = false;
      isTranslationMailActive = true;
    }

    this.state = {
      generatedText: "",
      generatedTextChinese: "",
      startText: "",
      startTextSave: "",
      finalMailText: "",
      isLoading: false,
      isGenerateBusinessMailActive: isGenerateBusinessMailActive,
      isSummarizeMailActive: isSummarizeMailActive,
      isTranslationMailActive: isTranslationMailActive,
      summary: "",
      translation: "",
    };
  }

  showGenerateBusinessMail = () => {
    this.setState({ isGenerateBusinessMailActive: true, isSummarizeMailActive: false, isTranslationMailActive: false });
  };

  showSummarizeMail = () => {
    this.setState({ isGenerateBusinessMailActive: false, isSummarizeMailActive: true, isTranslationMailActive: false });
  };

  showTranslateMail = () => {
    this.setState({ isGenerateBusinessMailActive: false, isSummarizeMailActive: false, isTranslationMailActive: true });
  };

  handleExpandClick = async () => {
    try {
      console.log("handleExpandClick1");
      const hf = new HfInference("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");
      //1
      let message1 =
        "The tower is 324 metres (1,063 ft) tall, about the same height as an 81-storey building, and the tallest structure in Paris. Its base is square, measuring 125 metres (410 ft) on each side. During its construction, the Eiffel Tower surpassed the Washington Monument to become the tallest man-made structure in the world, a title it held for 41 years until the Chrysler Building in New York City was finished in 1930.";
      // eslint-disable-next-line @typescript-eslint/no-unused-vars
      message1 =
        "在这个信息化的时代，技术的发展日新月异，给我们的生活带来了翻天覆地的变化。互联网的普及使得信息获取变得更加便捷，我们可以随时随地通过手机、电脑等设备浏览新闻、学习知识、与朋友交流。人工智能的进步则进一步改变了我们的生活方式，从智能家居到自动驾驶，科技正逐渐渗透到生活的各个角落。同时，电子商务的发展也使得购物变得更加方便快捷，网购已成为许多人日常生活的一部分。然而，伴随着技术的飞速发展，信息安全问题也日益突出，我们在享受便利的同时，也需提高自身的安全意识，保护好个人隐私。总之，技术的发展为我们带来了诸多便利，但也需要我们不断适应与防范。";
      // const resp = await hf.summarization({
      //   model: "facebook/bart-large-cnn",
      //   //model: "mistralai/Mistral-7B-Instruct-v0.2",
      //   inputs: message1,
      //   parameters: {
      //     max_length: 500,
      //   },
      // });
      //2
      // const res = await hf.textGeneration({
      //   model: "mistralai/Mistral-7B-Instruct-v0.2",
      //   inputs: "The answer to the universe is",
      // });
      //3
      // const res = await hf.chatCompletion({
      //   model: "mistralai/Mistral-7B-Instruct-v0.2",
      //   messages: [{ role: "user", content: "Complete the this sentence with words one plus one is equal " }],
      //   max_tokens: 500,
      //   temperature: 0.1,
      //   seed: 0,
      // });
      //4
      // const res = await hf.chatCompletion({
      //   model: "mistralai/Mistral-7B-Instruct-v0.2",
      //   messages: [
      //     {
      //       role: "user",
      //       content: "You are a helpful assistant that can help users to create professional business content.",
      //     },
      //     {
      //       role: "assistant",
      //       content: "ok.",
      //     },
      //     {
      //       role: "user",
      //       content:
      //         "Turn the following text into a professional business mail: " + "hello tom ,i will come back soon.",
      //     },
      //   ],
      //   max_tokens: 500,
      //   temperature: 0.1,
      //   seed: 0,
      // });
      //5
      const resp = await hf.chatCompletion({
        model: "mistralai/Mistral-7B-Instruct-v0.2",
        messages: [
          {
            role: "user",
            content: "can you recognize chinese?",
          },
          // {
          //   role: "assistant",
          //   content: "ok.",
          // },
          // {
          //   role: "user",
          //   content:
          //     "Turn the following text into a professional business mail: " + "hello tom ,i will come back soon.",
          // },
        ],
        max_tokens: 500,
        temperature: 0.1,
        seed: 0,
      });
      console.log("res", resp);
    } catch (error) {
      console.error("Error during API call", error);
    } finally {
      console.log("finally");
    }
  };

  //not used
  generateText0 = async () => {
    console.log("start in generate text00");
    const hf = new HfInference("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");
    var current = this;
    current.setState({ isLoading: true });
    const response = await hf.chatCompletion({
      model: "mistralai/Mistral-7B-Instruct-v0.2",
      messages: [
        {
          role: "user",
          content: "You are a helpful assistant that can help users to create content.",
        },
        {
          role: "assistant",
          content: "ok.",
        },
        {
          role: "user",
          content: "Turn the following text into a mail: " + this.state.startText,
        },
      ],
      max_tokens: 500,
      temperature: 0.1,
      seed: 0,
    });
    console.log("res", response.choices[0].message.content);
    current.setState({ isLoading: false });
    current.setState({ generatedText: response.choices[0].message.content });
    //current.setState({ generatedText: response});
  };

  generateText = async () => {
    // eslint-disable-next-line no-undef
    console.log("start in generate text123547");
    var current = this;
    if (current.state.startText.length === 0) {
      console.log("错误没有输入文本0");
      current.setState({ generatedText: "输入错误:没有输入文本!" });
      current.setState({ generatedTextChinese: "输入错误:没有输入文本!" });
      return;
    }
    if (current.state.startText === null) {
      console.log("错误没有输入文本1");
      current.setState({ generatedText: "输入错误:没有输入文本!!" });
      current.setState({ generatedTextChinese: "输入错误:没有输入文本!！" });
      return;
    }
    if (current.state.startText === "") {
      console.log("错误没有输入文本2");
      current.setState({ generatedText: "输入错误:没有输入文本!!!" });
      current.setState({ generatedTextChinese: "输入错误:没有输入文本!!！" });
      return;
    }
    console.log("starttext:" + "[" + current.state.startText + "]");
    console.log("startTextSave:" + "[" + current.state.startTextSave + "]");
    const configuration = new Configuration({
      //apiKey: "sk-proj-8yGEcGp2R0FxHt5d4Y0rT3BlbkFJXHuIQzkOCGCLkOoIlOvL",
      apiKey: "sk-proj-V0KmYBjbZDr9LsCdymd8T3BlbkFJ8tmYhizFiGltO0UmDkFV", //收费密钥
    });
    delete configuration.baseOptions.headers["User-Agent"];
    const openai = new OpenAIApi(configuration);
    current.setState({ isLoading: true });
    const response = await openai.createChatCompletion({
      model: "gpt-3.5-turbo",
      //model: "gpt-4",
      messages: [
        {
          role: "system",
          content: "You are a helpful assistant that can help users to create content.",
          //content: "You are a helpful assistant that can help users to create professional business content.",
        },
        {
          role: "user",
          content: "Turn the following text into a English mail.Do not be too loog.: " + current.state.startText,
        },
      ],
    });
    current.setState({ generatedText: response.data.choices[0].message.content });

    const messages1: ChatCompletionRequestMessage[] = [
      {
        role: "system",
        content: "You are a helpful assistant that can help users to translate content.",
        //content: "You are a helpful assistant that can help users to create professional business content.",
      },
      {
        role: "user",
        content:
          "Translate the following text into a Chinese mail.Do not be too loog.: " +
          response.data.choices[0].message.content,
        //content: "将下面的内容变成简体中文邮件，而且不要写太长.: " + current.state.startText,
      },
    ];
    const response1 = await openai.createChatCompletion({
      model: "gpt-3.5-turbo",
      messages: messages1,
    });
    console.log("chinese mail is" + response1.data.choices[0].message.content);
    current.setState({ generatedTextChinese: response1.data.choices[0].message.content });
    current.setState({ isLoading: false });
    current.setState({ startTextSave: current.state.startText }); // 神奇的bug 如果用this就会出现 第一次进入函数不生效，下次进入函数生效
    console.log("startTextSave after set:" + current.state.startTextSave);
    console.log("response:" + response.data.choices[0].message.content);
    console.log("generatedText", current.state.generatedText);
  };

  insertIntoMail = () => {
    const finalText = this.state.finalMailText.length === 0 ? this.state.generatedText : this.state.finalMailText;
    Office.context.mailbox.item.body.setSelectedDataAsync(finalText, {
      coercionType: Office.CoercionType.Text,
    });
  };
  insertIntoMailChinese = () => {
    const finalText = this.state.generatedTextChinese;
    Office.context.mailbox.item.body.setSelectedDataAsync(finalText, {
      coercionType: Office.CoercionType.Text,
    });
  };
  onTranslate = async () => {
    try {
      this.setState({ isLoading: true });
      const traslation = await this.translateMail0();
      this.setState({ translation: traslation, isLoading: false });
    } catch (error) {
      this.setState({ translation: error, isLoading: false });
    }
  };

  onSummarize = async () => {
    try {
      this.setState({ isLoading: true });
      const summary = await this.summarizeMail0();
      this.setState({ summary: summary, isLoading: false });
    } catch (error) {
      this.setState({ summary: error, isLoading: false });
    }
  };

  translateMail0(): Promise<any> {
    return new Office.Promise(function (resolve, reject) {
      try {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
          //const hf = new HfInference("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");
          const configuration = new Configuration({
            apiKey: "sk-proj-V0KmYBjbZDr9LsCdymd8T3BlbkFJ8tmYhizFiGltO0UmDkFV",
          });
          const openai = new OpenAIApi(configuration);

          let mailText = asyncResult.value.split(" ").slice(0, 1000).join(" "); //完整的邮件body，所有对话过程
          const senderName = Office.context.mailbox.item.from.displayName; //发件人名字
          const senderEmail = Office.context.mailbox.item.from.emailAddress; //发件地址
          let substr = senderName + " " + "<" + senderEmail + ">";
          let index = mailText.indexOf(substr);
          let submailtext = "";
          if (index !== -1) {
            submailtext = mailText.substring(0, index); //利用邮件人和邮件地址切割获取第一份邮件
          } else {
            submailtext = mailText;
          }
          substr = "发件人:";
          index = submailtext.lastIndexOf(substr);
          if (index !== -1) {
            submailtext = submailtext.substring(0, index); //去掉收件人：四个字符，最终这是最近一封邮件的内容
          }

          // eslint-disable-next-line @typescript-eslint/no-unused-vars
          const messages1: ChatCompletionRequestMessage[] = [
            {
              role: "system",
              content: "You are a helpful assistant that can translate text between languages.",
            },
            {
              role: "user",
              content: "Translate the following text into [Chinese]: " + submailtext,
            },
          ];
          const response1 = await openai.createChatCompletion({
            model: "gpt-3.5-turbo",
            messages: messages1,
          });
          let chinesecontent = response1.data.choices[0].message.content;
          let sumfinalres = "中文翻译:\n" + chinesecontent + "\n\n" + "original content:\n" + submailtext;
          //resolve(response.data.choices[0].message.content);
          //resolve(response1.data.choices[0].message.content);
          resolve(sumfinalres);
          //let mailtextaddsm = mailText + "senderName:[" + senderName + "] " + "senderEmail:[" + senderEmail + "]";
          //resolve("submailtext" + submailtext + "[" + mailtextaddsm + "]");
          //resolve(submailtext);
          //resolve(mailText);
          //resolve(submailtext);
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  summarizeMail0(): Promise<any> {
    return new Office.Promise(function (resolve, reject) {
      try {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
          //const hf = new HfInference("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");
          const configuration = new Configuration({
            apiKey: "sk-proj-V0KmYBjbZDr9LsCdymd8T3BlbkFJ8tmYhizFiGltO0UmDkFV",
          });
          const openai = new OpenAIApi(configuration);

          let mailText = asyncResult.value.split(" ").slice(0, 1000).join(" "); //完整的邮件body，所有对话过程
          const senderName = Office.context.mailbox.item.from.displayName; //发件人名字
          const senderEmail = Office.context.mailbox.item.from.emailAddress; //发件地址
          let substr = senderName + " " + "<" + senderEmail + ">";
          let index = mailText.indexOf(substr);
          let submailtext = "";
          if (index !== -1) {
            submailtext = mailText.substring(0, index); //利用邮件人和邮件地址切割获取第一份邮件
          } else {
            submailtext = mailText;
          }
          substr = "发件人:";
          index = submailtext.lastIndexOf(substr);
          if (index !== -1) {
            submailtext = submailtext.substring(0, index); //去掉收件人：四个字符，最终这是最近一封邮件的内容
          }
          const messages: ChatCompletionRequestMessage[] = [
            {
              role: "system",
              content:
                "You are a helpful assistant that can help users to better manage emails. The mail thread can be made by multiple prompts.",
            },
            {
              role: "user",
              content: "Only Summarize the following mail thread and summarize it with a bullet list: " + submailtext,
            },
          ];

          const response = await openai.createChatCompletion({
            model: "gpt-3.5-turbo",
            messages: messages,
          });
          // const response = await hf.summarization({
          //   model: "facebook/bart-large-cnn",
          //   //model: "mistralai/Mistral-7B-Instruct-v0.2",
          //   inputs: "how are you today,tom?you do not look good",
          //   parameters: {
          //     max_length: 500,
          //   },
          // });
          // const response = await hf.chatCompletion({
          //   model: "mistralai/Mistral-7B-Instruct-v0.2",
          //   messages: [
          //     {
          //       role: "user",
          //       content: "sum.",
          //     },
          //   ],
          //   max_tokens: 500,
          //   temperature: 0.1,
          //   seed: 0,
          // });
          console.log(response.data.choices[0].message.content);
          let summarymailres = response.data.choices[0].message.content;
          // eslint-disable-next-line @typescript-eslint/no-unused-vars
          const messages1: ChatCompletionRequestMessage[] = [
            {
              role: "system",
              content: "You are a helpful assistant that can translate text between languages.",
            },
            {
              role: "user",
              content: "Translate the following text into [Chinese]: " + summarymailres,
            },
          ];
          const response1 = await openai.createChatCompletion({
            model: "gpt-3.5-turbo",
            messages: messages1,
          });
          let summarymailreschinese = response1.data.choices[0].message.content;
          let sumfinalres =
            "中文总结:\n" + summarymailreschinese + "\n\n" + "original summarization:\n" + summarymailres;
          //resolve(response.data.choices[0].message.content);
          //resolve(response1.data.choices[0].message.content);
          resolve(sumfinalres);
          //let mailtextaddsm = mailText + "senderName:[" + senderName + "] " + "senderEmail:[" + senderEmail + "]";
          //resolve("submailtext" + submailtext + "[" + mailtextaddsm + "]");
          //resolve(submailtext);
          //resolve(mailText);
          //resolve(submailtext);
        });
      } catch (error) {
        reject(error);
      }
    });
  }
  summarizeMail(): Promise<any> {
    return new Office.Promise(function (resolve, reject) {
      try {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
          const configuration = new Configuration({
            apiKey: "sk-proj-V0KmYBjbZDr9LsCdymd8T3BlbkFJ8tmYhizFiGltO0UmDkFV",
          });
          const openai = new OpenAIApi(configuration);

          //take only the first 800 words of the mail
          //const mailText = asyncResult.value.split(" ").slice(0, 800).join(" ");
          let mailText = asyncResult.value.split(" ").slice(0, 1000).join(" "); //完整的邮件body，所有对话过程
          const senderName = Office.context.mailbox.item.from.displayName; //发件人名字
          const senderEmail = Office.context.mailbox.item.from.emailAddress; //发件地址
          let substr = senderName + " " + "<" + senderEmail + ">";
          let index = mailText.indexOf(substr);
          let submailtext = mailText.substring(0, index); //利用邮件人和邮件地址切割获取第一份邮件
          substr = "发件人:";
          index = submailtext.indexOf(substr);
          submailtext = submailtext.substring(0, index); //去掉收件人：四个字符，最终这是最近一封邮件的内容
          //create the request body
          const messages: ChatCompletionRequestMessage[] = [
            {
              role: "system",
              content:
                "You are a helpful assistant that can help users to better manage emails. The mail thread can be made by multiple prompts.",
            },
            {
              role: "user",
              content: "Summarize the following mail thread and summarize it with a bullet list: " + submailtext,
            },
          ];

          const response = await openai.createChatCompletion({
            model: "gpt-3.5-turbo",
            messages: messages,
          });
          resolve(submailtext);
          resolve(response.data.choices[0].message.content);
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  ProgressSection = () => {
    if (this.state.isLoading) {
      return <Progress title="Loading..." message="The AI is working..." />;
    } else {
      return <> </>;
    }
  };

  BusinessMailSection = () => {
    if (this.state.isGenerateBusinessMailActive) {
      return (
        <>
          <p>Briefly describe what you want to write in the mail:</p>
          <textarea
            className="ms-welcome"
            defaultValue={this.state.startTextSave}
            onChange={(e) => this.setState({ startText: e.target.value })}
            rows={5}
            cols={40}
          />
          <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.generateText}
            >
              Generate text
            </DefaultButton>
          </p>
          <this.ProgressSection />
          <textarea
            className="ms-welcome"
            value={this.state.generatedText}
            onChange={(e) => this.setState({ finalMailText: e.target.value })}
            rows={15}
            cols={40}
          />
          <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.insertIntoMail}
            >
              Insert into English mail
            </DefaultButton>
          </p>
          <p>Translation of email:</p>
          <textarea
            className="ms-welcome"
            value={this.state.generatedTextChinese}
            // onChange={(e) => this.setState({ finalMailText: e.target.value })}
            rows={10}
            cols={40}
          />
          {/* <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.insertIntoMailChinese}
            >
              Insert into Chinese mail
            </DefaultButton>
          </p> */}
        </>
      );
    } else {
      return <div> </div>;
    }
  };

  SummarizeMailSection = () => {
    if (this.state.isSummarizeMailActive) {
      return (
        <>
          <p>Summarize mail</p>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.onSummarize}
          >
            Summarize mail
          </DefaultButton>
          <this.ProgressSection />
          <textarea className="ms-welcome" defaultValue={this.state.summary} rows={15} cols={40} />
        </>
      );
    } else {
      return <div> </div>;
    }
  };

  translationMailSection = () => {
    if (this.state.isTranslationMailActive) {
      return (
        <>
          <p>Translate mail</p>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.onTranslate}
          >
            Translate mail
          </DefaultButton>
          <this.ProgressSection />
          <textarea className="ms-welcome" defaultValue={this.state.translation} rows={15} cols={40} />
        </>
      );
    } else {
      return <div> </div>;
    }
  };
  generateText1 = () => {
    try {
      console.log("Starting generateText function12");

      const apiKey = "sk-proj-8yGEcGp2R0FxHt5d4Y0rT3BlbkFJXHuIQzkOCGCLkOoIlOvL"; // Replace with your actual OpenAI API key
      console.log("4444444");
      const configuration = new Configuration({
        apiKey: apiKey,
      });
      delete configuration.baseOptions.headers["User-Agent"];
      console.log("33333");
      const openai = new OpenAIApi(configuration);
      console.log("11111");
      const response = openai.createChatCompletion({
        model: "gpt-3.5-turbo",
        messages: [
          { role: "system", content: "You are a helpful assistant." },
          { role: "user", content: "Turn the following text into a professional business mail: " }
        ],
      });
      console.log("11111222222");
      console.log("API Response:", response);
      console.log("response format:", response);
    } catch (error) {
      console.error("Error generating text:", error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <main className="ms-welcome__main">
          <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
            Outlook AI Assistant
          </h2>

          <p className="ms-font-l ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
            Choose your service:
          </p>
          <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.showGenerateBusinessMail}
            >
              Generate business mail
            </DefaultButton>
          </p>
          <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.showSummarizeMail}
            >
              Summarize mail
            </DefaultButton>
          </p>
          <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.showTranslateMail}
            >
              Translate mail
            </DefaultButton>
          </p>
          {/* <div>
            <DefaultButton onClick={this.handleExpandClick}>Generate Text for debug</DefaultButton>
          </div> */}

          <div>
            <this.BusinessMailSection />
          </div>
          <div>
            <this.SummarizeMailSection />
          </div>
          <div>
            <this.translationMailSection />
          </div>
        </main>
      </div>
    );
  }
}
