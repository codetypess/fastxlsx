import type { SetHyperlinkOptions } from "../types.js";
import {
  buildExternalHyperlinkXml,
  buildInternalHyperlinkXml,
  getHyperlinkRelationshipId,
  HYPERLINK_RELATIONSHIP_TYPE,
  removeHyperlinkFromSheetXml,
  upsertHyperlinkInSheetXml,
} from "./sheet-metadata.js";
import { getNextRelationshipIdFromXml, removeRelationshipById, upsertRelationship } from "./sheet-package.js";

export function setSheetHyperlink(
  sheetXml: string,
  relationshipsXml: string,
  address: string,
  target: string,
  options: SetHyperlinkOptions = {},
): { sheetXml: string; relationshipsXml: string } {
  const currentRelationshipId = getHyperlinkRelationshipId(sheetXml, address);
  let nextRelationshipsXml = relationshipsXml;

  if (target.startsWith("#")) {
    if (currentRelationshipId) {
      nextRelationshipsXml = removeRelationshipById(nextRelationshipsXml, currentRelationshipId);
    }

    return {
      sheetXml: upsertHyperlinkInSheetXml(
        sheetXml,
        buildInternalHyperlinkXml(address, target, options.tooltip),
        address,
      ),
      relationshipsXml: nextRelationshipsXml,
    };
  }

  const relationshipId = currentRelationshipId ?? getNextRelationshipIdFromXml(nextRelationshipsXml);
  nextRelationshipsXml = upsertRelationship(
    nextRelationshipsXml,
    relationshipId,
    HYPERLINK_RELATIONSHIP_TYPE,
    target,
    "External",
  );

  return {
    sheetXml: upsertHyperlinkInSheetXml(
      sheetXml,
      buildExternalHyperlinkXml(address, relationshipId, options.tooltip),
      address,
    ),
    relationshipsXml: nextRelationshipsXml,
  };
}

export function removeSheetHyperlink(
  sheetXml: string,
  relationshipsXml: string,
  address: string,
): { sheetXml: string; relationshipsXml: string } {
  const currentRelationshipId = getHyperlinkRelationshipId(sheetXml, address);
  return {
    sheetXml: removeHyperlinkFromSheetXml(sheetXml, address),
    relationshipsXml: currentRelationshipId
      ? removeRelationshipById(relationshipsXml, currentRelationshipId)
      : relationshipsXml,
  };
}
