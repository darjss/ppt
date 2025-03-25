import pptxgen from "pptxgenjs";
import fs from "fs";
import path from "path";
import https from "https";
import { fileURLToPath } from "url";

interface Slide {
  title: string;
  subtitle?: string;
  content?: string;
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
    imageSrc:
      "https://files.techhive.com/wp-content/uploads/2022/01/facebook_security_privacy_settings-100843288-orig.jpg",
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

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Create a new PowerPoint presentation
const pres = new pptxgen();

// Set presentation properties
pres.title = "DDoS Халдлагууд - Танилцуулга";
pres.subject = "Кибер аюулгүй байдал";
pres.author = "Same App";
pres.company = "Same";

// Define a theme for the presentation
pres.theme = {
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  background: { color: "0F172A" }, // dark blue
  titleColor: "FFFFFF",
  bodyColor: "E2E8F0",
};

// Define a master slide with the gradient background
pres.defineSlideMaster({
  title: "MASTER_SLIDE",
  background: { color: "0F172A" },
  objects: [
    {
      rectangle: {
        x: 0,
        y: 0,
        w: "100%",
        h: "100%",
        fill: {
          type: "gradient",
          color1: "0F172A", // dark blue
          color2: "1E293B", // lighter blue
          angle: 45,
        },
      },
    },
    {
      text: {
        text: "DDoS Халдлагууд - Танилцуулга",
        options: {
          x: 0.5,
          y: 6.8,
          w: "95%",
          h: 0.4,
          align: "center",
          fontSize: 10,
          color: "94A3B8",
          bold: false,
        },
      },
    },
  ],
});

// Function to download images
async function downloadImage(url: string, imagePath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    // Skip if file already exists
    if (fs.existsSync(imagePath)) {
      resolve(imagePath);
      return;
    }

    // Create directory if it doesn't exist
    const dir = path.dirname(imagePath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }

    // Download image
    const file = fs.createWriteStream(imagePath);
    https
      .get(url, (response) => {
        response.pipe(file);
        file.on("finish", () => {
          file.close();
          resolve(imagePath);
        });
      })
      .on("error", (err) => {
        fs.unlink(imagePath, () => {});
        reject(err);
      });
  });
}

// Function to generate slides
async function generatePresentation(): Promise<void> {
  try {
    // Create the images directory if it doesn't exist
    const imagesDir = path.resolve(process.cwd(), "images");
    if (!fs.existsSync(imagesDir)) {
      fs.mkdirSync(imagesDir, { recursive: true });
    }

    // Process each slide
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      if (slide === undefined) {
        throw new Error("Slide undefined");
      }
      console.log(`Processing slide ${i + 1}: ${slide.title}`);

      // Create a new slide
      const pptSlide = pres.addSlide({ masterName: "MASTER_SLIDE" });

      // Add title
      pptSlide.addText(slide.title, {
        x: 0.5,
        y: 0.5,
        w: "95%",
        h: 0.8,
        fontSize: 36,
        color: "FFFFFF",
        bold: true,
      });

      // Add subtitle if present
      if (slide.subtitle) {
        pptSlide.addText(slide.subtitle, {
          x: 0.5,
          y: 1.3,
          w: "95%",
          h: 0.5,
          fontSize: 24,
          color: "7DD3FC", // blue
          bold: false,
        });
      }

      // Handle slide content
      if (slide.content) {
        pptSlide.addText(slide.content, {
          x: 0.5,
          y: slide.subtitle ? 1.9 : 1.4,
          w: slide.imageSrc ? "45%" : "95%",
          h: 1.0,
          fontSize: 18,
          color: "E2E8F0", // light gray
          bold: false,
        });
      }

      // Add bullet points if present
      if (slide.bulletPoints && slide.bulletPoints.length > 0) {
        const bulletPointsY = slide.content
          ? slide.subtitle
            ? 3.0
            : 2.5
          : slide.subtitle
          ? 1.9
          : 1.4;

        pptSlide.addText(
          slide.bulletPoints.map((point) => ({ text: point, bullet: true })),
          {
            x: 0.5,
            y: bulletPointsY,
            w: slide.imageSrc ? "45%" : "95%",
            h: 3.0,
            fontSize: 16,
            color: "E2E8F0", // light gray
            bullet: { type: "bullet" },
          }
        );
      }

      // Add image if present
      if (slide.imageSrc) {
        try {
          // Format the image filename
          const imageExt = path.extname(slide.imageSrc) || ".png";
          const imageName = `slide_${i + 1}_image${imageExt}`;
          const imagePath = path.join(imagesDir, imageName);

          // Download image
          await downloadImage(slide.imageSrc, imagePath);

          // Add image to slide
          pptSlide.addImage({
            path: imagePath,
            x: slide.content || slide.bulletPoints ? "50%" : 0.5,
            y: 1.9,
            w: slide.content || slide.bulletPoints ? "48%" : "95%",
            h: 4.0,
            sizing: {
              type: "contain",
              w: slide.content || slide.bulletPoints ? "48%" : "95%",
              h: 4.0,
            },
          });
        } catch (err) {
          console.error(`Error adding image for slide ${i + 1}:`, err);
        }
      }

      // Add slide number
      pptSlide.addText(`${i + 1}/${slides.length}`, {
        x: 9.0,
        y: 6.5,
        w: 0.5,
        h: 0.3,
        fontSize: 12,
        color: "94A3B8", // gray
        bold: false,
      });
    }

    // Save the PowerPoint file
    const outputPath = path.join(process.cwd(), "lab4.pptx");
    await pres.writeFile({ fileName: outputPath });
    console.log(`Presentation saved to: ${outputPath}`);
  } catch (err) {
    console.error("Error generating presentation:", err);
  }
}

// Run the presentation generator
generatePresentation();
