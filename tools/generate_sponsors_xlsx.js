/**
 * sponsors.xlsx を初期データで(再)生成するスクリプト。
 * 通常は使わない(Excelファイルを直接編集すればよい)が、
 * 壊れた場合や列構成を変えたい場合の再生成用に残しておく。
 *
 * 実行前に一度だけ: npm install xlsx@0.18.5 --no-save
 * 実行: node generate_sponsors_xlsx.js
 */
const path = require('path');
const XLSX = require('xlsx');

const COLUMNS = [
  'id', 'group', 'group_label', 'tier', 'tier_label', 'order',
  'name', 'address', 'description', 'benefit',
  'website', 'website_label', 'instagram', 'line', 'images'
];

const COLUMN_NOTES = {
  id: 'ID(英数字・変更しない)',
  group: 'グループ(kobekatsu / topteam)',
  group_label: 'グループ表示名',
  tier: 'ティア(partner/main/silver/bronze/supplier)',
  tier_label: 'ティア表示名(空なら非表示)',
  order: '表示順(グループ・ティア内)',
  name: '企業・店舗名',
  address: '住所',
  description: '紹介文',
  benefit: 'EDUMIRAR会員特典',
  website: '公式サイト/店舗ページURL',
  website_label: 'サイトリンクのボタン文言(空なら「公式サイト」)',
  instagram: 'Instagramのユーザー名(@なし)',
  line: 'LINE等のURL(あれば)',
  images: '画像ファイル名(images/sponsors/<id>/配下、,区切り、ロゴを先頭に)'
};

const ROWS = [
  { id: 'remake', group: 'kobekatsu', group_label: 'コベカツクラブ', tier: 'partner', tier_label: '',
    order: 1, name: 'RE:MAKE', address: 'リモートで行います',
    description: '４ヶ月間のセミダイエットプログラムを提供しています！運動なしの食事管理のみのダイエットで平均-13%減量に成功中！無理のないダイエットをしたい方にオススメ！',
    benefit: '１ヶ月10000円→１ヶ月3500円(65%オフ)',
    website: '', website_label: '', instagram: 're_make_masashi', line: '', images: '01-logo.jpg' },

  { id: 'masaya-motoyama', group: 'kobekatsu', group_label: 'コベカツクラブ', tier: 'partner', tier_label: '',
    order: 2, name: '炭火焼鳥専門店まさやJR摂津本山店', address: '兵庫県神戸市東灘区岡本1-3-33',
    description: 'JR摂津本山駅北側出口を出て数秒！ご家族、お仕事終わりの方など幅広い層から人気の焼鳥屋です！団体様大歓迎！',
    benefit: 'ドリンク全品199円！(税込219円)(ボトルワイン除く)',
    website: 'https://s.tabelog.com/hyogo/A2801/A280105/28002781/', website_label: '食べログで見る',
    instagram: 'masaya.motoyama', line: '', images: '01.jpg' },

  { id: 'kobe-okamoto-seikotsuin', group: 'kobekatsu', group_label: 'コベカツクラブ', tier: 'partner', tier_label: '',
    order: 3, name: '神戸岡本鍼灸整骨院', address: '兵庫県神戸市東灘区岡本1丁目4-22オーガニックビル1F',
    description: 'JR摂津本山駅、阪急岡本駅から徒歩3分！型にハマった施術ではなく、その人のニーズに合わせた本格整骨院！',
    benefit: '※ご利用の際はチャットまでご連絡下さい。',
    website: 'https://kobeokamoto-seikotsuin.com', website_label: '公式サイト',
    instagram: '', line: '', images: '01-logo.jpg' },

  { id: 'laughome', group: 'topteam', group_label: 'トップチームスポンサー', tier: 'main', tier_label: 'MAIN SPONSOR',
    order: 1, name: '株式会社Laughomeラホーム不動産', address: '大阪府大阪市中央区東心斎橋1丁目18-10YERビル4F',
    description: '物件だけでなく、「安心」と「感動」を届ける不動産屋です！お客様のニーズに寄り添ったサポートを、「借りたい」「買いたい」「売りたい」「貸したい」というニーズにあわせてサポート致します！ご納得頂けるまで、徹底的に寄り添います。',
    benefit: 'ご成約でAmazonギフト券プレゼント！(賃貸5000円、売買50000円)',
    website: 'https://laughome.laugh-group.jp', website_label: '公式サイト',
    instagram: 'laughome_fudousan', line: '', images: '01-logo.jpg,02.jpg' },

  { id: 'bar-laughtale', group: 'topteam', group_label: 'トップチームスポンサー', tier: 'silver', tier_label: 'SILVER SPONSOR',
    order: 1, name: 'BAR Laugh Tale', address: '大阪市東心斎橋2-5-23ふぁみーゆ笠屋町3F',
    description: '若きオーナーが経営しているミナミのBARです！しっぽりよりみんなで楽しくワイワイできるフランクなBARで楽しめちゃう空間！',
    benefit: '※ご利用の際はチャットまでご連絡ください',
    website: '', website_label: '', instagram: 'bar_laughtale0211', line: '', images: '01-logo.jpg' },

  { id: 'yakitori-smith', group: 'topteam', group_label: 'トップチームスポンサー', tier: 'silver', tier_label: 'SILVER SPONSOR',
    order: 2, name: '焼鳥スミス', address: '大阪府大阪市北区天神橋5丁目5-28シエルブルー天神橋102',
    description: '焼鳥もドリンクもボリューム満点でコスパ抜群！お仕事帰りの宴会でもご利用頂けます！オーナーこだわりの焼鳥を召し上がってください！',
    benefit: '※ご利用の際はチャットまでご連絡ください',
    website: 'https://s.tabelog.com/osaka/A2701/A270103/27140617/', website_label: '食べログで見る',
    instagram: 'yakitorismith', line: '', images: '01-logo.jpg' },

  { id: 'neofirst-osaka', group: 'topteam', group_label: 'トップチームスポンサー', tier: 'bronze', tier_label: 'BRONZE SPONSOR',
    order: 1, name: 'Neo First 大阪福島店', address: '大阪府大阪市福島区福島５丁目５−６ 福島テラス 402',
    description: 'JR福島駅から徒歩30秒！大阪北区で口コミランキングTOP3に入る本格エステが体験できます！チケット制ではなく都度払いにも関わらずこの価格でできるエステは見た事がありません！',
    benefit: '※特典利用の際はチャットまで連絡ください',
    website: '', website_label: '', instagram: 'neofirst_osaka', line: '', images: '01.jpg,02.jpg' },

  { id: 'ohara-seikotsuin', group: 'topteam', group_label: 'トップチームスポンサー', tier: 'bronze', tier_label: 'BRONZE SPONSOR',
    order: 2, name: '小原整骨院', address: '大阪市阿倍野区昭和町2-1-7',
    description: 'EDUMIRAR大阪TOP選手がいつもお世話になっている接骨院です！スポーツ障害はもちろん、様々なケアに対応して下さる接骨院です。',
    benefit: '※ご利用の際はチャットまでご連絡ください',
    website: 'https://ohara-seikotsuin.jp/access/', website_label: '公式サイト',
    instagram: '', line: '', images: '01-logo.jpg' },

  { id: 'yamaguchi-cosmo-dental', group: 'topteam', group_label: 'トップチームスポンサー', tier: 'bronze', tier_label: 'BRONZE SPONSOR',
    order: 3, name: '山口コスモデンタルクリニック', address: '大阪府大阪市住之江区南港北1丁目31-17リバーガーデンコスモスクエア',
    description: 'EDUMIRAR大阪の親父こと、背番号20山口剛選手が医院長の歯医者さんです！幅広く歯の治療を行なっていただけます！',
    benefit: '※ご利用の際はチャットまでご連絡ください',
    website: 'http://www.y-cosmo-dent.com/', website_label: '公式サイト',
    instagram: '', line: '', images: '01-logo.jpg' },

  { id: 'masaya-ashiya', group: 'topteam', group_label: 'トップチームスポンサー', tier: 'bronze', tier_label: 'BRONZE SPONSOR',
    order: 4, name: '焼鳥まさやJR芦屋店', address: '兵庫県芦屋市大原町2-6ラモール芦屋104',
    description: 'JR芦屋駅近くで30年以上営業している焼き鳥屋です！普段使いからみんなの集まりで、老若男女問わずお気軽に楽しんでいただけます！',
    benefit: '公式LINEから会員証提示でお会計から10%引き',
    website: 'https://s.tabelog.com/hyogo/A2803/A280302/28001308/', website_label: '食べログで見る',
    instagram: 'masaya_ashiyahigashi', line: '', images: '01.jpg,02.jpg' },

  { id: 'gobernante', group: 'topteam', group_label: 'トップチームスポンサー', tier: 'supplier', tier_label: 'SUPPLIER',
    order: 1, name: 'GOBERNANTE sport.', address: 'ECサイト',
    description: 'EDUMIRAR大阪TOPチームのサプライヤーブランド様です！元Fリーガー田中優輝さんデザインのアパレルブランドです。',
    benefit: '※ご利用の際はチャットまでご連絡ください',
    website: 'https://gobernante.official.ec', website_label: 'ECサイトで見る',
    instagram: 'gobernante_official', line: '', images: '01-logo.jpg' }
];

function build() {
  const wb = XLSX.utils.book_new();

  const aoa = [COLUMNS, ...ROWS.map(row => COLUMNS.map(col => row[col] !== undefined ? row[col] : ''))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = COLUMNS.map(col => ({ wch: col === 'description' ? 50 : (col === 'benefit' || col === 'website' ? 32 : 20) }));
  ws['!freeze'] = { xSplit: 0, ySplit: 1 };
  XLSX.utils.book_append_sheet(wb, ws, 'sponsors');

  const guideAoa = [
    ['列名', '説明'],
    ['このファイルについて',
      'EDUMIRARスポンサー紹介LPの企業情報(特典・紹介文・リンク等)を管理するファイルです。' +
      'このファイルを編集して保存し、サイトの「コーディング/data/sponsors.xlsx」を置き換えれば、' +
      'サイト側は自動的に最新の内容を表示します(サイトのHTML/JSファイルは触らなくてOKです)。'],
    ['行の追加・削除', '新しい協賛企業を追加する場合は1行追加してidから入力してください。削除も行ごとで構いません。'],
    ...COLUMNS.map(col => [col, COLUMN_NOTES[col]])
  ];
  const guideWs = XLSX.utils.aoa_to_sheet(guideAoa);
  guideWs['!cols'] = [{ wch: 28 }, { wch: 80 }];
  XLSX.utils.book_append_sheet(wb, guideWs, '使い方');

  const outPath = path.join(__dirname, '..', 'data', 'sponsors.xlsx');
  XLSX.writeFile(wb, outPath);
  console.log('saved:', outPath);
}

build();
