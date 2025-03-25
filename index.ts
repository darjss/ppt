import pptxgen from "pptxgenjs";

// Define the Slide interface
interface Slide {
  title: string;
  subtitle?: string;
  content: string;
  bulletPoints?: string[];
  imageSrc?: string;
  imageAlt?: string;
}


export const slides: Slide[] = [
    {
      title: "Компьютерын аюулгүй байдал",
      subtitle:
        "Зөвшөөрөлгүй нэвтрэлт, вэбкам аюулгүй байдал, нийгмийн сүлжээний вирусын эсрэг хамгаалалт",
      content:
        "Энэхүү танилцуулгад бид компьютерын аюулгүй байдлын гурван чухал сэдвийг авч үзэх болно: зөвшөөрөлгүй компьютерт нэвтрэхээс сэргийлэх арга, вэбкамераа хамгаалах, мөн Facebook-ээс вирус авах эрсдэлээс хэрхэн сэргийлэх талаар.",
      imageSrc:
        "https://www.cloudflare.com/img/learning/security/threats/malware/virus-vs-malware.svg",
      imageAlt: "Компьютерын аюулгүй байдлын ерөнхий зураглал",
    },
    {
      title: "Зөвшөөрөлгүй компьютерт нэвтрэхээс хэрхэн сэргийлэх вэ?",
      content:
        "Зөвшөөрөлгүй компьютерт нэвтрэх нь таны хувийн мэдээлэл, санхүүгийн өгөгдөлд аюул учруулж болзошгүй.",
      bulletPoints: [
        "Хүчтэй, өвөрмөц нууц үг ашиглах",
        "Хоёр алхамт баталгаажуулалт (2FA) идэвхжүүлэх",
        "Үйлдлийн системээ байнга шинэчилж байх",
        "Найдвартай хорт кодын эсрэг програм суулгах",
        "Гал хана (файрволл) идэвхжүүлэх",
        "Хэрэглээгүй үед компьютераа түгжих",
      ],
      imageSrc:
        "https://cybernews.com/wp-content/uploads/2023/06/Two-factor-authentication-security-1.png",
      imageAlt: "Хоёр алхамт баталгаажуулалтын зураглал",
    },
    {
      title: "Хүчтэй нууц үг үүсгэх аргууд",
      content:
        "Хүчтэй нууц үг нь таны компьютерын аюулгүй байдлын эхний шугам юм.",
      bulletPoints: [
        "Дор хаяж 12 тэмдэгт бүхий урт нууц үг ашиглах",
        "Том, жижиг үсэг, тоо болон тусгай тэмдэгтүүдийг хослуулах",
        "Өөр өөр үйлчилгээнд өөр нууц үг ашиглах",
        "Хувийн мэдээлэл агуулаагүй байх (нэр, төрсөн өдөр гэх мэт)",
        "Нууц үг менежер ашиглах (LastPass, Bitwarden гэх мэт)",
        "Нууц үгээ тогтмол солих (3-6 сар тутамд)",
      ],
      imageSrc:
        "https://www.security.org/wp-content/uploads/2019/12/strong-password-display.jpg",
      imageAlt: "Хүчтэй нууц үгийн жишээ",
    },
    {
      title: "Хоёр алхамт баталгаажуулалт (2FA)",
      content:
        "Хоёр алхамт баталгаажуулалт нь таны аккаунтыг хамгаалах нэмэлт хамгаалалтын давхарга юм.",
      bulletPoints: [
        "Нууц үгээс гадна нэмэлт хамгаалалтын давхарга үүсгэдэг",
        "Бүртгэл рүү нэвтрэхэд хоёр дахь алхам шаарддаг (утасны код, апп-ийн батлах код гэх мэт)",
        "Таны нууц үг алдагдсан ч аккаунтыг хамгаалдаг",
        "SMS-ээр баталгаажуулах, аппликейшн-аар баталгаажуулах, биометрик гэх мэт олон төрөлтэй",
        "Нэвтрэх эрх мэдээллийг хулгайлах халдлагаас хамгаална",
        "Аль болох олон үйлчилгээ дээр идэвхжүүлэх нь зүйтэй",
      ],
      imageSrc:
        "https://www.okta.com/sites/default/files/styles/tinypng/public/media/image/2020-02/Authentication%20type%20image.png",
      imageAlt: "Хоёр алхамт баталгаажуулалтын төрлүүд",
    },
    {
      title: "Вэбкамераа хакердуулахаас хэрхэн хамгаалах вэ?",
      content:
        "Таны вэбкамерыг зөвшөөрөлгүйгээр удирдах нь хувийн орон зайд халдах ноцтой зөрчил юм.",
      bulletPoints: [
        "Ашиглаагүй үед вэбкамаа физик хаалт ашиглан хаах",
        "Найдвартай антивирус програм хангамж суулгах",
        "Вэбкамын драйверуудыг байнга шинэчлэх",
        "Вэбкамын зөвшөөрлийг хянах (аль программууд ашиглаж болохыг тохируулах)",
        "Сэжигтэй вебсайт, аппуудаас зайлсхийх",
        "Интернэт холболтын аюулгүй байдлыг хангах",
      ],
      imageSrc:
        "https://www.makeuseof.com/wp-content/uploads/2020/06/webcam-cover-3.jpg",
      imageAlt: "Вэбкамын физик хаалт",
    },
    {
      title: "Вэбкамын зөвшөөрлийг хянах",
      content:
        "Аль нэг програм таны вэбкамерыг ашиглаж байгааг хянах нь чухал юм.",
      bulletPoints: [
        "Windows дээр: Тохиргоо > Нууцлал > Камер сонголтоор орж зөвшөөрлийг удирдах",
        "Mac дээр: System Preferences > Security & Privacy > Camera хэсэгт хандах",
        "Вэбкамын идэвхтэй байдлыг харуулах индикаторыг анзаарах (ихэнхдээ гэрэл)",
        "Зөвхөн шаардлагатай үед л программуудад камер ашиглах зөвшөөрөл өгөх",
        "Таньд хэрэггүй аппликейшнуудын камер хандалтыг хаах",
        "Камерын ашиглалтыг хянах нэмэлт хяналтын програмууд ашиглах",
      ],
      imageSrc:
        "https://piunikaweb.com/wp-content/uploads/2022/01/Windows-11-webcam-permission.jpg",
      imageAlt: "Windows-д камерын зөвшөөрлийг удирдах",
    },
    {
      title: "Facebook-ээс вирус авах боломжтой юу?",
      content:
        "Тийм, Facebook нь вирус болон бусад төрлийн хортой програм тараах платформ болж болзошгүй.",
      bulletPoints: [
        "Сэжигтэй линк, хавсралтаар дамжин тархдаг",
        "Хуурамч тоглоом, аппликейшнуудаар дамжин тархдаг",
        "Фишинг халдлага (хуурамч нэвтрэх хуудас)-аар дамжин",
        "Найзын дүр эсгэн хуурамч профайлаар дамжин",
        "Вирусжсан зураг, видеоны хуваалцаж болзошгүй",
        "Cookie хулгайлах скриптүүдээр дамжин",
      ],
      imageSrc:
        "https://us.norton.com/content/dam/blogs/images/norton/am/phishing-in-the-mobile-world.png",
      imageAlt: "Фишинг халдлагын жишээ",
    },
    {
      title: "Facebook-д аюулгүй байх аргууд",
      content:
        "Facebook ашиглахдаа дараах аюулгүй байдлын дадлуудыг баримтлах хэрэгтэй:",
      bulletPoints: [
        "Facebook аккаунтдаа хүчтэй нууц үг ашиглах",
        "Хоёр алхамт баталгаажуулалтыг идэвхжүүлэх",
        "Танихгүй хүмүүсийн найзын хүсэлтийг хүлээж авахгүй байх",
        "Сэжигтэй линк дээр дарахаас зайлсхийх",
        "Хачин танилцуулга, таны дуртай зүйлс, сонирхолтой саналууд амлаж буй контентоос болгоомжлох",
        "Фэйсбүүк аппликейшнаа шинэчлэгдсэн байлгах",
        "Нууцлалын тохиргоогоо тогтмол шалгаж, шинэчлэх",
      ],
      imageSrc:
        "https://scontent.fuln1-2.fna.fbcdn.net/v/t39.2365-6/120973513_338686077291355_8148888128带9984872_n.png?_nc_cat=104&ccb=1-7&_nc_sid=ad8a9d&_nc_ohc=vY_6IfY3lhkAX-FWXvZ&_nc_ht=scontent.fuln1-2.fna&oh=00_AfAgGcwEf3IVhq8WVYS5nF8XPYUzTiQKXxivb_EsjsJwDQ&oe=63B15C3D",
      imageAlt: "Facebook аюулгүй байдлын тохиргоо",
    },
    {
      title: "Нийгмийн сүлжээнд хамгийн түгээмэл тархдаг аюулууд",
      content:
        "Facebook болон бусад нийгмийн сүлжээнд дараах аюулууд түгээмэл тархдаг:",
      bulletPoints: [
        "Фишинг халдлага - хуурамч нэвтрэх хуудсууд таны мэдээллийг хулгайлдаг",
        "Скам - хуурамч бэлэг, мөнгөн шагнал, азтан болсон гэх худал мэдээлэл",
        "Нийгмийн инженерчлэл - сэтгэл хөдлөлд тулгуурласан манипуляци",
        "Хортой программ бүхий линкүүд - таны төхөөрөмжид вирус суулгадаг",
        "Хуурамч аппууд - хортой код бүхий тоглоом, квиз, хэрэгсэл",
        "Хуурамч профайлууд - таны найзын дүр эсгэн хуурамч аккаунтаар халдах",
      ],
    //   imageSrc:
    //     "https://files.techhive.com/wp-content/uploads/2022/01/facebook_security_privacy_settings-100843288-orig.jpg",
      imageAlt: "Нийгмийн сүлжээний аюулгүй байдлын эрсдэлүүд",
    },
    {
      title: "Нийтлэг вирусны халдварын шинж тэмдгүүд",
      content:
        "Компьютерт вирус халдварласан эсэхийг эдгээр шинж тэмдгээр мэдэж болно:",
      bulletPoints: [
        "Компьютер гэнэт удааширч, гацах",
        "Өөрөө өөрийгөө дахин эхлүүлэх",
        "Таны зөвшөөрөлгүйгээр файлууд нээгдэх, хаагдах",
        "Диск зай ойлгомжгүй шалтгаанаар дүүрэх",
        "И-мэйл нь таны эзэмшигчийн мэдэлгүйгээр илгээгдэх",
        "Антивирус програм гэнэт ажиллахаа болих",
        "Хачин попап цонхнууд гарч ирэх",
        "Интернэт браузерын нүүр хуудас өөрчлөгдөх",
      ],
      imageSrc:
        "https://www.avg.com/content/dam/avg/images/slider/slow-computer-causes-shutterstock-624276236.png",
      imageAlt: "Компьютерын удаан ажиллагаа - вирусын нэг шинж тэмдэг",
    },
    {
      title: "Хортой кодоос хамгаалах аргууд",
      content:
        "Вирус болон бусад хортой кодоос компьютероо хамгаалах үндсэн аргууд:",
      bulletPoints: [
        "Найдвартай антивирус програм суулгаж, байнга шинэчлэх",
        "Үйлдлийн систем болон бусад программуудыг тогтмол шинэчлэх",
        "Сэжигтэй и-мэйл хавсралтуудыг нээхгүй байх",
        "Итгэлтэй эх сурвалжаас програм татаж суулгах",
        "Олон нийтийн Wi-Fi-д VPN ашиглах",
        "Чухал файлуудаа тогтмол нөөцлөх",
        "Windows файрволл идэвхжүүлэх",
        "Web браузерын аюулгүй байдлын тохиргоог шалгах",
      ],
      imageSrc:
        "https://www.researchgate.net/publication/347497505/figure/fig15/AS:969588881207309@1608010232142/Types-of-computer-virus.png",
      imageAlt: "Компьютерын вирусын төрлүүд",
    },
    {
      title: "Аюулгүй интернэт хэрэглээний зөвлөмжүүд",
      content:
        "Интернэтэд аюулгүй байхын тулд дараах зөвлөмжүүдийг баримтлах хэрэгтэй:",
      bulletPoints: [
        "HTTPS (нууцлагдсан холболт) бүхий вебсайтуудыг ашиглах",
        "Браузерын өргөтгөлүүдээ хязгаарлах, зөвхөн шаардлагатайг нь ашиглах",
        "Хувийн мэдээллээ хуваалцахдаа болгоомжтой байх",
        "Бүртгэл үүсгэхдээ сайтын нууцлалын бодлогыг уншиж танилцах",
        "Шинэ сайтууд дээр танигдсан нууц үгээ ашиглахгүй байх",
        "Файл татахаас өмнө эх сурвалжийг шалгах",
        "Криптжуулсан холболт (VPN) ашиглах",
        "Хувийн мэдээллийг нийгмийн сүлжээнд хязгаарлах",
      ],
      imageSrc: "https://www.apa.org/images/web-security_tcm7-262475.jpg",
      imageAlt: "Аюулгүй интернэт хэрэглээ",
    },
    {
      title: "Компьютерын аюулгүй байдлын шилдэг туршлагууд",
      content:
        "Компьютерын аюулгүй байдлын ерөнхий шилдэг туршлагуудыг дагаж мөрдөх нь чухал:",
      bulletPoints: [
        "Бүх төхөөрөмж дээрээ найдвартай хамгаалалтын програм суулгах",
        "Програм хангамжаа байнга шинэчлэх (автомат шинэчлэлтийг идэвхжүүлэх)",
        "Чухал өгөгдлөө тогтмол нөөцлөх (3-2-1 нөөцлөлийн зарчим)",
        "Аюулгүй байдлын талаар мэдлэгээ дээшлүүлэх",
        "Олон нийтийн Wi-Fi сүлжээнүүдийг ашиглахдаа болгоомжтой байх",
        "Бүх төхөөрөмждөө хүчтэй нууц үг ашиглах",
        "Биометрик баталгаажуулалтыг боломжтой бол ашиглах",
        "Шинэ төрлийн кибер аюул заналхийллийн талаар мэдээлэлтэй байх",
      ],
      imageSrc:
        "https://www.kaspersky.com/content/en-global/images/repository/isc/2020/what-is-cyber-security.jpg",
      imageAlt: "Компьютерын аюулгүй байдлын шилдэг туршлагууд",
    },
    {
      title: "Мэдлэг шалгах асуултууд",
      content: "Өөрийн мэдлэгээ шалгахын тулд дараах асуултууд:",
      bulletPoints: [
        "Хоёр алхамт баталгаажуулалт гэж юу вэ?",
        "Таны бодлоор хамгийн аюултай вирусын халдварын шинж тэмдэг аль нь вэ?",
        "Facebook дээр сэжигтэй линк таарвал та юу хийх вэ?",
        "Таны вэбкамерыг хакерууд яаж ашиглаж болох вэ?",
        "Гар утас, компьютер дээрх хүчтэй нууц үгийн ялгаа юу вэ?",
        "Аюулгүй харилцаа холбооны хувьд та ямар арга хэрэгсэл ашигладаг вэ?",
      ],
      imageSrc:
        "https://www.researchgate.net/publication/322393233/figure/fig1/AS:669952597303305@1536739874564/IT-security-measures-for-prevention-reaction-evaluation-and-awareness.png",
      imageAlt: "Аюулгүй байдлын арга хэмжээний төрлүүд",
    },
    {
      title: "Дүгнэлт",
      content:
        "Компьютерын аюулгүй байдал нь зөвхөн тусгайлсан мэдлэгтэй хүмүүст биш, бидний хүн нэг бүрт хамааралтай чухал сэдэв юм.",
      bulletPoints: [
        "Зөвшөөрөлгүй нэвтрэлтээс сэргийлэхийн тулд хүчтэй нууц үг, хоёр алхамт баталгаажуулалт ашиглах",
        "Вэбкамын аюулгүй байдлыг хангахын тулд физик хаалт ашиглах, зөвшөөрлийг хянах",
        "Facebook болон бусад нийгмийн сүлжээнээс вирус авахаас сэргийлэхийн тулд сэжигтэй линк, хавсралтаас зайлсхийх",
        "Компьютерын аюулгүй байдал нь тасралтгүй үйл явц бөгөөд байнгын сонор сэрэмжтэй байдлыг шаарддаг",
        "Өөрийн мэдээллийн хамгаалалт нь таны хариуцлага юм",
      ],
      imageSrc:
        "https://cdn.pixabay.com/photo/2017/01/17/12/43/security-1986451_1280.jpg",
      imageAlt: "Компьютерын аюулгүй байдлын дүгнэлт",
    },
  ];

// Create a new presentation
const createPresentation = async () => {
  try {
    const pres = new pptxgen();

    // Set presentation properties
    pres.layout = "LAYOUT_16x9";
    pres.author = "PptxGenJS";
    pres.title = "Computer Security Presentation";
    pres.subject = "Computer Security";

    // Define common slide styles
    const titleStyle = {
      fontFace: "Arial",
      fontSize: 36,
      color: "363636",
      bold: true,
    };

    const subtitleStyle = {
      fontFace: "Arial",
      fontSize: 20,
      color: "666666",
      italic: true,
    };

    const contentStyle = {
      fontFace: "Arial",
      fontSize: 16,
      color: "333333",
    };

    // Create slides
    for (const slide of slides) {
      const newSlide = pres.addSlide();

      // Add title
      newSlide.addText(slide.title, {
        ...titleStyle,
        x: 0.5,
        y: 0.5,
        w: "90%",
        h: 0.8,
      });

      // Add subtitle if exists
      if (slide.subtitle) {
        newSlide.addText(slide.subtitle, {
          ...subtitleStyle,
          x: 0.5,
          y: 1.4,
          w: "90%",
          h: 0.6,
        });
      }

      // Add content
      const contentY = slide.subtitle ? 2.1 : 1.5;
      newSlide.addText(slide.content, {
        ...contentStyle,
        x: 0.5,
        y: contentY,
        w: slide.imageSrc ? "55%" : "90%",
        h: 1,
      });

      // Add bullet points if exists - adding each bullet point separately
      if (slide.bulletPoints && slide.bulletPoints.length > 0) {
        const bulletY = contentY + 1.2;
        
        // Create a text object for each bullet point instead of passing an array
        for (let i = 0; i < slide.bulletPoints.length; i++) {
          newSlide.addText(slide.bulletPoints[i], {
            fontFace: "Arial",
            fontSize: 16,
            color: "333333",
            bullet: true,
            x: 0.5,
            y: bulletY + (i * 0.4), // Adjust vertical position for each bullet
            w: slide.imageSrc ? "55%" : "90%",
            h: 0.4,
          });
        }
      }

      // Add image if exists
      if (slide.imageSrc) {
        try {
          newSlide.addImage({
            path: slide.imageSrc,
            x: "60%",
            y: contentY,
            w: "35%",
            h: 3,
            altText: slide.imageAlt || slide.title,
          });
        } catch (error) {
          console.error(`Failed to add image for slide "${slide.title}":`, error);
          
          // Add a placeholder text instead of the image
          newSlide.addText("Image could not be loaded", {
            fontFace: "Arial",
            fontSize: 14,
            color: "999999",
            italic: true,
            x: "60%",
            y: contentY + 1.5,
            w: "35%",
            h: 0.5,
            align: "center",
          });
        }
      }
    }

    // Save the presentation
    await pres.writeFile({ fileName: "computer_security_presentation.pptx" });
    console.log("Presentation created successfully!");
  } catch (error) {
    console.error("Error creating presentation:", error);
  }
};

// Run the function
createPresentation();
