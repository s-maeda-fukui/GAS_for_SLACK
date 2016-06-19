/*
 * Googleフォーム自動返信スクリプト
 */

/**
 * 起点・設定
 */
function sendMailForm() {
    Logger.log('sendMailForm() debug start');

    //------------------------------------------------------------
    // 設定エリアここから
    //------------------------------------------------------------

    // 件名、本文、フッター
    var subject = "七夕祭当日スタッフへのご登録完了のお知らせ";

    //body　初期値
    var body = "七夕祭当日スタッフにご登録いただいた皆様\n\n"
    body += "この度は、第１１回七夕祭当日スタッフにご登録いただき誠にありがとうございます。\ｎ";
    body += "以下の内容でご登録完了しましたのでお知らせ致します。\nご確認お願い致します。\n";

    var footer = "当日の流れ等詳細につきましては後日改めてご連絡致します。\n";
    footer += "また、やむを得ない理由でキャンセルされる場合は七夕祭開催一週間前までに下記メールアドレスまで必ずご連絡をよろしくお願い致します。\n\n";
    footer += "七夕祭の成功は皆様にかかっています！\n"
    footer += "どうぞよろしくお願い致します。\n"
    footer += "何かご質問等ございましたら、下記メールアドレスまでお問い合わせください。\n\n";
    footer += "-- \n";
    footer += "神戸大学 六甲台学生評議会(法・経済・経営学部ゼミ幹事会)\n"
    footer += "B.E.L. Student Council\n";
    footer += "bel.tanabata2016@gmail.com"

    var add_tmp = "";

    // 入力カラム名の指定
    var NAME_COL_NAME = '名前（フルネーム）';
    var NAME_DEP = '学部(神戸大学以外の方は大学名をご記入ください)';
    var MAIL_COL_NAME = 'メールアドレス';
    var NUM_PHONE = "電話番号";
    var PART_1 = '担当したい役割(第１希望)';
    var PART_2 = '担当したい役割(第２希望)';
    var PART_3 = '担当したい役割(第３希望)';
    var CHARGE_TIME = '参加可能時間';

    // メール除外カラム
    var EXCLUDE_COLS = ['ステータス','対応日時'];
    var stmp ;

    // メール送信先
    var admin = "s.maeda.kobe@gmail.com"; // 管理者（必須、ユーザーメールのReplyTo、および管理者メールのTOになります）

    //------------------------------------------------------------
    // 設定エリアここまで
    //------------------------------------------------------------

    try{
        // スプレッドシートの操作
        var book = SpreadsheetApp.getActive();
        var sh   = book.getSheetByName("DB");
        //シートの最終行を取得
        var rows = sh.getLastRow();
        //シートの最終列を取得
        var cols = sh.getLastColumn();
        //データが入力されている範囲を獲得
        var rg   = sh.getDataRange();
        //ログに書き込み
        Logger.log("rows="+rows+" cols="+cols);


        var to = "";    // To: （入力者のアドレスが自動で入ります）

        /* メール件名・本文作成と送信先メールアドレス取得 */
        //jは列指定（入力のある）最終列まで回す
        for ( var j = 1; j <= cols-2; j++ ) {
            var col_name  = rg.getCell(1, j).getValue();    // カラム名（最上位行を参照）
            // rows：最終行が格納されている
            var col_value = rg.getCell(rows, j).getValue(); // 入力値

            // メール用変換
            if ( col_name === NAME_COL_NAME ) {
                //格納値が”名前”のとき
                body += "【"+col_name+"】\n";
                body += col_value+" 様\n";

            }else if ( col_name === MAIL_COL_NAME ) {
                //格納値が”メールアドレス”のとき
                to = col_value;

            }else if( colname === NAME_DEP ){
                //格納値が'学部(神戸大学以外の方は大学名をご記入ください)
                body += "【"+col_name+"】\n";
                body += col_value+"\n";
            }else if( colname === NUM_PHONE ){
                //格納値が"電話番号"
                body += "【"+col_name+"】\n";
                body += col_value+"\n";
            }else if( colname === PART_1 ){
                //格納値が'担当したい役割(第１希望)'
                body += "【"+col_name+"】\n";
                body += col_value+"\n";
            }else if( colname === PART_2 ){
                //格納値が'担当したい役割(第２希望)'
                body += "【"+col_name+"】\n";
                body += col_value+"\n";
            }else if( colname === PART_3 ){
                //格納値が'担当したい役割(第３希望)'
                body += "【"+col_name+"】\n";
                body += col_value+"\n";
            }else if( colname === CHARGE_TIME ){
                //格納値が'参加可能時間'
                add_tmp += "【"+col_name+"】\n";
                add_tmp += col_value+"\n";
            }else{
              //DO NOTHING
            }//if メール用変換

            // 日付フォーマットの変換
            // 他にも変換したいカラムがある場合はこのif分をコピーしてカラム名・日付フォーマットを変更する
            if ( col_name === 'タイムスタンプ' ) {
                col_value = Utilities.formatDate(col_value, "JST", "yyyy-MM-dd HH:mm:ss");
                //タイムスタンプをコピー
                stmp = col_value ;
            }

            // メール送信除外カラム
            if ( EXCLUDE_COLS.length > 0 ) {
                var is_exclude = false;
                //EXCLUDE_COLS.length should be 2
                for ( var k = 0; k < EXCLUDE_COLS.length; k++ ) {
                    if ( col_name === EXCLUDE_COLS[k] ) {
                        // === は同値演算子
                        is_exclude = true;
                        break;
                    }//if
                }//for

                if ( is_exclude ) {
                    // 除外カラムなのでスキップ
                    continue;
                }//if

            }//if

        }

        //最終連結
        body += add_tmp;
        //フッターの連結
        body += footer;

        // メール送信
        if ( to ) {
            //フォーム登録者に送信
            MailApp.sendEmail(to, subject, body);

            sh.getRange(rows, cols-3).setValue("DONE");
            sh.getRange(rows, cols-2).setValue(stmp);

            // 連続で送るとエラーになるので1秒スリープ
            Utilities.sleep(1000);

            //管理者にも同様のメールを送信
            MailApp.sendEmail(admin, subject, body);

        }else{
            MailApp.sendEmail(admin, "【失敗】Googleフォームにメールアドレスが指定されていません", body);
        }//if

        // 連続で送るとエラーになるので1秒スリープ
        Utilities.sleep(1000);

    }catch(e){
        MailApp.sendEmail(admin, "【失敗】Googleフォームからメール送信中にエラーが発生", e.message);
    }//try文（例外発生用）

}//sendMail

