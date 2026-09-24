/**
 * 단위: 저장 검증에 쓰는 XML 적격성 검사 (src/XmlWellFormed.ts).
 *
 * 0.3.3 의 verify_integrity 는 글자 모양 세 가지(`<?xml` 유무, 끝이 `<` 로 끊김,
 * 태그 안의 `<`)만 봤다. 닫는 태그가 어긋난 section 은 셋 다 통과해 "검증 완료"로
 * 나갔다 (회신 2026-09-24 ②: 저장본 <hp:p> 5730 열림 / 5728 닫힘).
 */
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { xmlWellFormednessError, findMalformedXmlParts } from '../../src/XmlWellFormed';

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>';
const NS = 'xmlns:hs="urn:hs" xmlns:hp="urn:hp"';

describe('xmlWellFormednessError', () => {
  it('정상 문서는 null', () => {
    expect(xmlWellFormednessError(`${HEAD}<hs:sec ${NS}><hp:p><hp:run><hp:t>가</hp:t></hp:run></hp:p></hs:sec>`)).toBeNull();
  });

  it('닫는 태그가 하나 빠진 문서를 잡는다 — 0.3.3 검사 세 가지를 모두 통과하던 모양', () => {
    const xml = `${HEAD}<hs:sec ${NS}><hp:p><hp:run><hp:t>가</hp:t></hp:run></hs:sec>`;
    // 0.3.3 의 세 검사가 이 문서를 통과시켰다는 것을 같이 고정한다.
    expect(xml.includes('<?xml')).toBe(true);
    expect(/<[^>]*$/.test(xml)).toBe(false);
    expect(/<[^>]*</.test(xml)).toBe(false);
    expect(xmlWellFormednessError(xml)).toMatch(/close tag/i);
  });

  it('닫는 태그 이름이 어긋난 문서를 잡는다', () => {
    expect(xmlWellFormednessError(`${HEAD}<hs:sec ${NS}><hp:p><hp:run></hp:p></hp:run></hs:sec>`)).not.toBeNull();
  });

  it('글 안의 맨 & 를 잡는다 (escape 누락)', () => {
    expect(xmlWellFormednessError(`${HEAD}<hs:sec ${NS}><hp:t>A & B</hp:t></hs:sec>`)).not.toBeNull();
  });

  it('escape 된 글과 CDATA 는 통과한다', () => {
    expect(xmlWellFormednessError(`${HEAD}<hs:sec ${NS}><hp:t>A &amp; B &lt;C&gt;</hp:t><hp:t><![CDATA[<x>]]></hp:t></hs:sec>`)).toBeNull();
  });
});

describe('findMalformedXmlParts', () => {
  it('깨진 부분만 경로와 함께 돌려주고, XML 이 아닌 부분은 보지 않는다', async () => {
    const zip = new JSZip();
    zip.file('mimetype', 'application/hwp+zip');
    zip.file('Contents/content.hpf', `${HEAD}<opf:package xmlns:opf="urn:opf"/>`);
    zip.file('Contents/section0.xml', `${HEAD}<hs:sec ${NS}><hp:p></hs:sec>`);
    zip.file('Contents/section1.xml', `${HEAD}<hs:sec ${NS}><hp:p/></hs:sec>`);
    zip.file('BinData/image1.png', '<not xml');

    const bad = await findMalformedXmlParts(zip);
    expect(bad).toHaveLength(1);
    expect(bad[0]).toMatch(/^Contents\/section0\.xml: /);
  });
});
