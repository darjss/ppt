import pptxgen from "pptxgenjs";
import { fetch } from "bun";
import { writeFileSync, mkdirSync, existsSync } from "fs";
import { join } from "path";

// Define the Slide interface
interface Slide {
  title: string;
  subtitle?: string;
  content: string;
  bulletPoints?: string[];
  imageSrc?: string;
  imageAlt?: string;
}

// Your slide data

export const slides: Slide[] = [
  {
    title: "Кредит карт хулгайлах болон залилан",
    subtitle: "Цахим гэмт хэргийн нэгэн түгээмэл төрөл",
    content: "Интернетээр дамжуулан хийгддэг санхүүгийн залилангийн хамгийн түгээмэл хэлбэрүүдийн нэг"
  },
  {
    title: "Агуулга",
    content: "Энэхүү танилцуулгад дараах сэдвүүдийг хөндөнө:",
    bulletPoints: [
      "Кредит карт хулгайлалт ба залилангийн тодорхойлолт",
      "Энэ төрлийн гэмт хэргийн хэлбэрүүд",
      "Гэмт этгээдүүдийн ашигладаг аргууд",
      "Бодит кэйс судалгаа",
      "Дэлхий дахины тоо баримт",
      "Монгол дахь нөхцөл байдал",
      "Өөрийгөө хамгаалах аргууд",
      "Банкны хамгаалалтын системүүд",
      "Хууль эрх зүйн орчин",
      "Ирээдүйн чиг хандлага",
      "Дүгнэлт"
    ]
  },
  {
    title: "Кредит карт хулгайлалт ба залилангийн тодорхойлолт",
    content: "Кредит карт хулгайлалт ба залилан гэдэг нь бусдын кредит карт эсвэл дебит картын мэдээллийг зөвшөөрөлгүйгээр олж авч, санхүүгийн ашиг олох зорилгоор ашиглах үйлдлийг хэлнэ.",
    bulletPoints: [
      "Карт эзэмшигчийн зөвшөөрөлгүйгээр хувийн мэдээллийг олж авах",
      "Картын мэдээллийг ашиглан хууль бус худалдан авалт хийх",
      "Картын мэдээллийг бусдад зарах, солилцох"
    ]
  },
  {
    title: "Кредит карт хулгайлалт ба залилангийн түүх",
    content: "Кредит картын залилан дижитал эрин үеэс өмнө ч оршин байсан:",
    bulletPoints: [
      "1960-аад он: Анхны кредит картууд хэрэглээнд нэвтэрсэн",
      "1970-аад он: Кредит картын залилангийн анхны тохиолдлууд бүртгэгдэж эхэлсэн",
      "1980-90-ээд он: Картын дугаарыг хуулж хуурамчаар үйлдвэрлэх аргууд түгээмэл болсон",
      "2000-аад он: Онлайн худалдаа өсөхийн хэрээр цахим залилан нэмэгдсэн",
      "2010-аад он: Том хэмжээний өгөгдлийн зөрчлүүд (data breaches) түгээмэл болсон",
      "2020-оод он: Криптовалют, NFT зэрэг шинэ технологид суурилсан залилангууд гарч ирсэн"
    ]
  },
  {
    title: "Кредит карт хулгайлалт ба залилангийн хэлбэрүүд",
    content: "Кредит картын залилан нь олон төрлийн аргаар хийгддэг:",
    bulletPoints: [
      "Онлайн (CNP) залилан - Карт байхгүй үед хийгддэг залилан (Card Not Present)",
      "Скиммер ашиглах - ATM болон POS терминалд скиммер төхөөрөмж суурилуулж, картын мэдээлэл хулгайлах",
      "Фишинг халдлага - Хуурамч имэйл, вэбсайт ашиглан хэрэглэгчдийн мэдээллийг авах",
      "Карт клондох - Хуурамч карт үйлдвэрлэх"
    ]
  },
  {
    title: "Онлайн (CNP) залилан дэлгэрэнгүй",
    subtitle: "Card Not Present Fraud",
    content: "Онлайн залилан нь өнөөдөр хамгийн түгээмэл тохиолддог хэлбэр юм:",
    bulletPoints: [
      "Картын эзэмшигч биечлэн байхгүй тохиолдолд хийгддэг залилан",
      "Онлайн худалдаа, утсаар захиалга, имэйлээр гүйлгээ хийх үед гардаг",
      "Кибер гэмт хэрэгтнүүд хулгайлсан картын мэдээллийг ашиглан зөвшөөрөлгүй худалдан авалт хийдэг",
      "Ихэвчлэн дижитал бүтээгдэхүүн (тоглоом, аппликэйшн, дижитал тасалбар) худалдан авснаар залилангийн эсрэг хамгаалалтыг шалгадаг",
      "EMV чип технологи (смарт карт) нэвтэрсэнээр биечлэн хийгдэх залилан буурч, харин онлайн залилан нэмэгдсэн"
    ]
  },
  {
    title: "Скиммер ашиглах залилан дэлгэрэнгүй",
    content: "Скиммер нь картын мэдээлэл хулгайлахад зориулагдсан төхөөрөмж юм:",
    bulletPoints: [
      "ATM машин болон POS терминалд суурилуулдаг жижиг, нууц төхөөрөмж",
      "Хэрэглэгч картаа уншуулахад соронзон зурвасны мэдээллийг хуулбарладаг",
      "Ихэвчлэн жижиг камер эсвэл хуурамч гарын товчлуур нэмж суурилуулж ПИН кодыг хулгайлдаг",
      "Орчин үеийн скиммерууд Bluetooth технологи ашиглаж, алсаас мэдээлэл татах боломжтой",
      "Скиммер илрүүлэгч аппуудыг хэрэглэгчид ашиглах боломжтой",
      "Хулгайлагдсан мэдээллийг ашиглан хуурамч карт үйлдвэрлэх эсвэл онлайн залиланд ашигладаг"
    ]
  },
  {
    title: "Фишинг халдлагын дэлгэрэнгүй",
    content: "Фишинг нь хэрэглэгчдийг хуурамч сайт руу чиглүүлж мэдээлэл авах арга юм:",
    bulletPoints: [
      "Банк, карт гаргагч байгууллагын албан ёсны мэйл мэт харагдах имэйл илгээдэг",
      "Яаралтай асуудал үүссэн, аккаунт баталгаажуулах шаардлагатай гэх мэссэж агуулдаг",
      "Хуурамч вэбсайт руу чиглэсэн холбоос агуулдаг",
      "Хуурамч сайт нь албан ёсны сайттай адилхан харагддаг",
      "Хэрэглэгч нэр, нууц үг, картын мэдээлэл, нийгмийн даатгалын дугаар зэргийг оруулахыг шаарддаг",
      "Бизнес имэйл залилан (BEC) нь фишингийн нэг хэлбэр бөгөөд компанийн захирал, санхүүгийн албаны дарга нарын имэйлийг дуурайж, мөнгө шилжүүлэх хүсэлт илгээдэг"
    ]
  },
  {
    title: "Гэмт этгээдүүдийн ашигладаг аргууд",
    content: "Кредит картын залилан үйлдэгчид дараах аргуудыг түгээмэл ашигладаг:",
    bulletPoints: [
      "Техникийн аргууд: Өгөгдлийн зөрчил (Data breaches), вэбсайтын хакердах, малвэр/хортой код ашиглах, скиммер төхөөрөмж суурилуулах",
      "Нийгмийн инженерчлэлийн аргууд: Фишинг имэйл, мессежүүд, хуурамч утасны дуудлага (Vishing), найдвартай байгууллага, хүмүүсийн дүр эсгэх"
    ]
  },
  {
    title: "Нийгмийн инженерчлэлийн аргууд дэлгэрэнгүй",
    content: "Нийгмийн инженерчлэл нь хүмүүсийг сэтгэл зүйн аргаар удирдан мэдээлэл өгүүлэх арга юм:",
    bulletPoints: [
      "Итгэл төрүүлэх: Өөрийгөө банкны ажилтан, техникийн дэмжлэгийн ажилтан гэх мэтээр танилцуулах",
      "Яаралтай байдал: Хэрэглэгчийг яарч шийдвэр гаргахад хүргэх ('Таны дансыг хаах гэж байна', 'Яаралтай арга хэмжээ авах шаардлагатай')",
      "Айлган сүрдүүлэх: Аккаунт блоклогдох, торгууль төлөх зэргээр айлган сүрдүүлэх",
      "Хязгаарлагдмал боломж: 'Зөвхөн өнөөдөр л энэ үнэ хүчинтэй' гэх мэтээр шахалт үүсгэх",
      "Vishing (Voice phishing): Утсаар залилан хийх",
      "Smishing (SMS phishing): Мессежээр залилан хийх"
    ]
  },
  {
    title: "Бодит кэйс судалгаа: Таргет компанийн өгөгдлийн зөрчил (2013)",
    content: "2013 оны 11-12-р сард АНУ-ын \"Target\" сүлжээ дэлгүүрт томоохон өгөгдлийн зөрчил гарсан:",
    bulletPoints: [
      "40 сая гаруй үйлчлүүлэгчдийн кредит/дебит картын мэдээлэл алдагдсан",
      "70 сая үйлчлүүлэгчдийн хувийн мэдээлэл (нэр, хаяг, имэйл, утас) алдагдсан",
      "Хакерууд POS терминалуудаас мэдээлэл хулгайлах малвэр суулгасан",
      "Алдагдсан картын мэдээллүүдийг даркнэт дээр зарж, хууль бус худалдан авалт хийхэд ашигласан",
      "Target компани $18.5 сая доллар торгууль төлж, $202 сая доллар нөхөн төлбөр төлсөн"
    ]
  },
  {
    title: "Бодит кэйс: Equifax өгөгдлийн зөрчил (2017)",
    content: "2017 онд кредит скорингийн агентлаг Equifax томоохон мэдээллийн алдагдалд өртсөн:",
    bulletPoints: [
      "147 сая америкчуудын хувийн мэдээлэл алдагдсан",
      "Нийгмийн даатгалын дугаар, төрсөн он сар өдөр, хаягууд, зарим тохиолдолд жолооны үнэмлэхийн дугаар алдагдсан",
      "Хакерууд Apache Struts вэб аппликэйшны хүрээнд байсан аюулгүй байдлын цоорхойг ашигласан",
      "Equifax компани эмзэг байдлын талаар мэдэж байсан боловч засаагүй",
      "Equifax $700 сая долларын торгууль төлсөн",
      "Энэ нь түүхэн дэх хамгийн том хувийн мэдээллийн зөрчлүүдийн нэг болсон"
    ]
  },
  {
    title: "Бодит кэйс: Marriott зочид буудлын өгөгдлийн зөрчил (2018)",
    content: "2018 онд Marriott International зочид буудлын сүлжээ томоохон мэдээллийн алдагдалд өртсөн:",
    bulletPoints: [
      "500 сая хүртэлх зочдын мэдээлэл алдагдсан",
      "Нэр, хаяг, утасны дугаар, имэйл хаяг, паспортын дугаар, Marriott Rewards гишүүнчлэлийн мэдээлэл",
      "Зарим тохиолдолд кредит картын мэдээлэл алдагдсан",
      "Хакерууд 2014 оноос эхлэн системд нэвтэрсэн байж болзошгүй",
      "Энэ нь зочид буудлын салбарт тохиолдсон хамгийн том мэдээллийн алдагдал болсон",
      "Marriott $123 сая еврогийн торгууль төлсөн"
    ]
  },
  {
    title: "Дэлхий дахины тоо баримт",
    content: "Кредит картын залиланд холбоотой статистик мэдээлэл:",
    bulletPoints: [
      "2023 онд дэлхий даяар кредит картын залилангийн улмаас $38.5 тэрбум долларын хохирол учирсан",
      "Хэрэглэгчдийн 47% нь ямар нэгэн хэлбэрийн картын залиланд өртсөн туршлагатай",
      "Онлайн залилан (CNP) нь нийт залилангийн 81%-ийг эзэлдэг",
      "Залилангийн гүйлгээний дундаж хэмжээ нь $1,088 доллар",
      "Үйлчлүүлэгчдийн 40% нь картын залиланд өртсөний дараа тухайн карт гаргагч байгууллагыг сольдог",
      "EMV чипний нэвтрэлтийн дараа биечлэн (POS) хийгдэх залилан 80%-иар буурсан боловч онлайн залилан 40%-иар өссөн"
    ]
  },
  {
    title: "Монгол дахь нөхцөл байдал",
    content: "Монгол Улсад кредит картын залилангийн байдал:",
    bulletPoints: [
      "Монгол Улсад жилд дунджаар 500-600 орчим банкны картын залилангийн хэрэг бүртгэгддэг",
      "Хамгийн түгээмэл нь хуурамч вэбсайт, фишинг халдлага",
      "Монголбанкны мэдээллээр онлайн худалдаанд картаа ашиглах хандлага өсөхийн хэрээр залиланд өртох эрсдэл нэмэгдэж байна",
      "Банкууд 3D Secure найдвартай баталгаажуулалтын систем нэвтрүүлснээр залилангийн тохиолдол буурч байна",
      "Иргэдийн дунд банкны картын аюулгүй байдлын талаарх мэдлэг дутмаг байгаа нь залилангийн нэг шалтгаан болдог",
      "Гэмт хэргийн хуульд 2020 оноос цахим гэмт хэрэгтэй тэмцэх тусгай зүйл ангиуд нэмэгдсэн"
    ]
  },
  {
    title: "Өөрийгөө хамгаалах аргууд",
    content: "Кредит картын залилангаас хамгаалахын тулд дараах аргуудыг хэрэгжүүлэх нь чухал:",
    bulletPoints: [
      "Картын мэдээллээ хамгаалах: Найдвартай сайтуудад л өгөх, олон нийтийн Wi-Fi ашиглан санхүүгийн гүйлгээ хийхгүй байх, шифрлэлтэй холболтоор дамжуулах (https://)",
      "Тогтмол хяналт тавих: Картын хуулгаа байнга шалгах, мэдээллийн хяналтын үйлчилгээ ашиглах",
      "Аюулгүй байдлын дадал: Хүчтэй нууц үгс ашиглах, хоёр үе шаттай баталгаажуулалт ашиглах, картын мэдээллээ нийгмийн сүлжээнд хуваалцахгүй байх"
    ]
  },
  {
    title: "Банкны хамгаалалтын системүүд",
    content: "Банкууд картын залилангаас хамгаалах олон төрлийн технологи ашигладаг:",
    bulletPoints: [
      "EMV чип технологи: Соронзон зурвастай харьцуулахад илүү аюулгүй, карт клондоход хүндрэлтэй",
      "3D Secure: Онлайн худалдан авалтын үед нэмэлт баталгаажуулалт шаарддаг (SMS код, апп дахь баталгаажуулалт)",
      "Залилангийн илрүүлэлтийн систем: Алгоритмууд ашиглан сэжигтэй, хэвийн бус гүйлгээг илрүүлдэг",
      "Токенизаци: Жинхэнэ картын дугаарын оронд түр зуурын токен үүсгэх",
      "Геолокаци баталгаажуулалт: Хэрэглэгчийн байршил болон гүйлгээний байршлыг харьцуулан шалгах",
      "Гүйлгээний дүнгийн хязгаар: Тодорхой хэмжээнээс дээш дүнтэй гүйлгээнд нэмэлт баталгаажуулалт шаардах"
    ]
  },
  {
    title: "Хууль эрх зүйн орчин",
    content: "Кредит картын залилантай тэмцэх хууль эрх зүйн зохицуулалт:",
    bulletPoints: [
      "АНУ-ын Fair Credit Billing Act: Картын эзэмшигчид зөвшөөрөлгүй гүйлгээний хариуцлагыг $50-д хязгаарладаг",
      "Европын GDPR: Хувийн мэдээллийн хамгаалалт, өгөгдлийн зөрчлийн мэдэгдэх үүргийг тодорхойлдог",
      "PCI DSS: Карт гүйлгээ хийдэг бүх байгууллагууд дагаж мөрдөх ёстой аюулгүй байдлын стандарт",
      "Монгол Улсын Гэмт хэргийн тухай хууль: Цахим залилангийн гэмт хэргийг тусгайлан зохицуулдаг",
      "Монголбанкны журмууд: Банкны карт гаргах, гүйлгээ хийх үйл ажиллагааны аюулгүй байдлын шаардлагууд",
      "Банк, санхүүгийн байгууллагын харилцагчийн мэдээллийн нууцлалын тухай хууль"
    ]
  },
  {
    title: "Ирээдүйн чиг хандлага",
    content: "Картын залилантай тэмцэх ирээдүйн технологи ба чиг хандлага:",
    bulletPoints: [
      "Биометрийн баталгаажуулалт: Хурууны хээ, царайн таних технологи, нүдний торлогоор таних",
      "Хиймэл оюун ухаан: Залилангийн илрүүлэлтийн системүүд улам боловсронгуй болж, өөрөө суралцах алгоритмууд ашиглах",
      "Блокчейн технологи: Гүйлгээний найдвартай, өөрчлөх боломжгүй бүртгэл үүсгэх",
      "Мобайл төлбөрийн системүүд: Физик картад найдах хэрэгцээг бууруулснаар картын мэдээлэл хулгайлах боломжийг хязгаарлах",
      "Эцсийг хүртэл шифрлэлт: Мэдээллийг дамжуулах, хадгалах бүх шатанд шифрлэх",
      "Зам доторх төлөлт (in-path payment): Хэрэглэгчийн картын мэдээллийг худалдагч нарт ил болгохгүйгээр төлбөр хийх боломж"
    ]
  },
  {
    title: "Дүгнэлт",
    content: "Кредит картын хулгайлалт ба залилан нь цахим орчны аюулгүй байдлын нэгэн томоохон сорилт хэвээр байна:",
    bulletPoints: [
      "Кредит картын хулгайлалт ба залилан нь цахим гэмт хэргийн түгээмэл төрөл юм",
      "Технологи хөгжихийн хэрээр гэмт этгээдүүдийн арга улам нарийсаж байна",
      "Мэдлэгтэй байж, урьдчилан сэргийлэх арга хэмжээг авснаар эрсдэлээ бууруулж болно",
      "Хэрэглэгч, банк, хууль сахиулах байгууллагууд хамтран ажиллаж цогц аюулгүй байдлыг хангах шаардлагатай",
      "Технологийн хөгжил, аюулгүй байдлын шинэ системүүд нь хязгааргүй боломжуудыг авчирдаг боловч шинэ эрсдэлүүдийг ч бий болгодог",
      "Санхүүгийн байгууллагууд болон хуулийн байгууллагууд энэхүү асуудалтай тэмцэхэд хамтран ажиллах шаардлагатай"
    ]
  },
  {
    title: "Ашигласан эх сурвалж",
    content: "Энэхүү танилцуулгыг бэлтгэхэд дараах эх сурвалжуудыг ашигласан:",
    bulletPoints: [
      "https://www.techopedia.com/definition/26587/internetcrime",
      "https://www.ftc.gov/news-events/topics/identity-theft/credit-card-fraud",
      "https://www.cnbc.com/2019/01/23/heres-how-your-personal-data-gets-stolen-in-a-data-breach.html",
      "https://corporate.target.com/press/releases/2014/12/target-provides-update-on-data-breach-and-financia",
      "https://www.investopedia.com/financial-edge/0212/how-to-protect-yourself-from-credit-card-fraud.aspx",
      "https://www.mongodb.com/blog/post/marriott-data-breach",
      "https://www.bankrate.com/finance/credit-cards/credit-card-fraud-statistics/",
      "https://www.mongolbank.mn/documents/regulation/cards/index.html",
      "https://www.visa.com/blogarchives/us/2019/02/13/emv-chip-cards-helped-reduce-counterfeit-fraud-by-80-percent/index.html"
    ]
  }
];

// Function to download an image from a URL
async function downloadImage(url: string, index: number): Promise<string | null> {
  try {
    // Create images directory if it doesn't exist
    const imagesDir = join(process.cwd(), "images");
    if (!existsSync(imagesDir)) {
      mkdirSync(imagesDir);
    }

    // Generate a filename based on the index
    const fileExtension = url.split(".").pop()?.split("?")[0] || "png";
    const filename = join(imagesDir, `image_${index}.${fileExtension}`);

    // Download the image
    const response = await fetch(url);
    
    if (!response.ok) {
      console.error(`Failed to download image from ${url}: ${response.statusText}`);
      return null;
    }
    
    const imageBuffer = await response.arrayBuffer();
    writeFileSync(filename, Buffer.from(imageBuffer));
    
    console.log(`Downloaded image from ${url} to ${filename}`);
    return filename;
  } catch (error) {
    console.error(`Error downloading image from ${url}:`, error);
    return null;
  }
}

// Create a new presentation
const createPresentation = async () => {
  try {
    const pres = new pptxgen();

    // Set presentation properties
    pres.layout = "LAYOUT_16x9";
    pres.author = "PptxGenJS";
    pres.title = "Computer Security Presentation";
    pres.subject = "Computer Security";

    // Define dark blue theme colors
    const themeColors = {
      background: "0F2A4A", // Dark blue background
      title: "FFFFFF",      // White text for titles
      subtitle: "BDD6F5",   // Light blue for subtitles
      content: "FFFFFF",    // White text for content
      accent: "4A89DC",     // Accent blue for highlights
      bulletColor: "BDD6F5" // Light blue for bullet points
    };

    // Define common slide styles with the theme colors
    const titleStyle = {
      fontFace: "Arial",
      fontSize: 36,
      color: themeColors.title,
      bold: true,
    };

    const subtitleStyle = {
      fontFace: "Arial",
      fontSize: 20,
      color: themeColors.subtitle,
      italic: true,
    };

    const contentStyle = {
      fontFace: "Arial",
      fontSize: 16,
      color: themeColors.content,
    };

    // Create slides
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      const newSlide = pres.addSlide();

      if(slide===undefined){
        throw new Error("Slide is undefined");
      }
      // Set slide background color
      newSlide.background = { color: themeColors.background };

      // Add title with a subtle accent line below
      newSlide.addText(slide.title, {
        ...titleStyle,
        x: 0.5,
        y: 0.5,
        w: "90%",
        h: 0.8,
      });

      // Add accent line below title
      newSlide.addShape(pres.ShapeType.rect, {
        x: 0.5,
        y: 1.3,
        w: 3,
        h: 0.05,
        fill: { color: themeColors.accent },
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
        
        // Create a text object for each bullet point
        for (let j = 0; j < slide.bulletPoints.length; j++) {
          newSlide.addText(slide.bulletPoints[j], {
            fontFace: "Arial",
            fontSize: 16,
            color: themeColors.bulletColor,
            bullet: { type: "bullet", color: themeColors.accent },
            x: 0.5,
            y: bulletY + (j * 0.4), // Adjust vertical position for each bullet
            w: slide.imageSrc ? "55%" : "90%",
            h: 0.4,
          });
        }
      }

      // Add image if exists - download it first
      if (slide.imageSrc) {
        try {
          // Download the image
          const localImagePath = await downloadImage(slide.imageSrc, i);
          
          if (localImagePath) {
            newSlide.addImage({
              path: localImagePath,
              x: "60%",
              y: contentY,
              w: "35%",
              h: 3,
              altText: slide.imageAlt || slide.title,
            });
          } else {
            // Add a placeholder if image download failed
            newSlide.addText("Image could not be loaded", {
              fontFace: "Arial",
              fontSize: 14,
              color: themeColors.subtitle,
              italic: true,
              x: "60%",
              y: contentY + 1.5,
              w: "35%",
              h: 0.5,
              align: "center",
            });
          }
        } catch (error) {
          console.error(`Failed to add image for slide "${slide.title}":`, error);
          
          // Add a placeholder text instead of the image
          newSlide.addText("Image could not be loaded", {
            fontFace: "Arial",
            fontSize: 14,
            color: themeColors.subtitle,
            italic: true,
            x: "60%",
            y: contentY + 1.5,
            w: "35%",
            h: 0.5,
            align: "center",
          });
        }
      }

      // Add slide number
      newSlide.addText(`${i + 1}/${slides.length}`, {
        fontFace: "Arial",
        fontSize: 10,
        color: themeColors.subtitle,
        x: 9.0,
        y: 6.8,
        w: 0.5,
        h: 0.3,
        align: "right",
      });
    }

    // Save the presentation
    await pres.writeFile({ fileName: "lab7.pptx" });
    console.log("Presentation created successfully!");
  } catch (error) {
    console.error("Error creating presentation:", error);
  }
};

// Run the function
createPresentation();
